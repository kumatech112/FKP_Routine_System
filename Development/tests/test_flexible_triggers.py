"""All-workflow flexible intake and the reported XX/missing-extension/read-mail regression."""
from datetime import timedelta
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import tempfile,unittest,zipfile
import pyzipper
from routine.triggers import match_mail,names,infer_sources
from routine.matching import locate_body_file
from routine.outlook import Outlook
from routine.settings import default_settings,runtime_settings
from routine.workflows import FileWorkflowService,run
from routine.vendor.fkp.domain import ValidationError
from routine.vendor.fkp.persistence import BANGKOK,now
import threading

class ClassifyTests(unittest.TestCase):
    def test_actual_document_subject_typo_xx_and_missing_extension_uses_body_round_date(self):
        found=match_mail('customer_documents','เอกสารลูกค้าจาก E-Mail กลาง รอบ2 PM_2902026','A10xxL_Document_Round2_29102026',received_date='261029')
        self.assertEqual(found['file_names'],['A1015L_Document_Round2_29102026'])
        self.assertEqual((found['round'],found['mail_date']),('2','29102026'))
        self.assertNotIn('trigger_error',found)

    def test_every_task_accepts_filename_with_wrong_subject(self):
        cases={'customer_documents':'A10XXL_Document_Round2_29102026',
            'payment_online':'261029_A10XX_A10XX_TSS_PAYMENT_Online2.zip.gpg',
            'performance':'A10XX_Daily_Recovery_Performance_Oct26.zip.gpg',
            'performance_success_rate':'A10XX_Performance_SAM_OA_Oct26.zip.gpg',
            'enforcement':'A10XXL_Approval_Enforcement_BAY_29102026_2.pdf.pgp'}
        for task,name in cases.items():
            for source in ['body','attachment','subject']:
                with self.subTest(task=task,source=source):
                    result=match_mail(task,name if source=='subject' else 'หัวข้อผิด',name if source=='body' else '',[name] if source=='attachment' else [],'261029')
                    self.assertIsNotNone(result)
                    self.assertNotIn('trigger_error',result,result)
                    self.assertTrue(result['file_names'])

    def test_enforcement_without_date_can_still_read_identity_from_inline_image(self):
        found=match_mail("enforcement","แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์","")
        self.assertIsNotNone(found)
        self.assertNotIn("trigger_error",found)
        self.assertTrue(found["needs_file_identity"])

    def test_old_daily_mail_cannot_select_a_new_transport_file_by_month_alone(self):
        from routine.files import find_archive
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/"261029_A1015_A1015_Daily_Recovery_Performance_Oct26.zip.gpg";path.touch()
            scan={"network_folder":temp,"network_file_regex":default_settings()["workflows"]["performance"]["network_file_regex"]}
            match={"month_name":"Oct","year":"26","mail_date":"21102026","received_date":"261022","file_names":["A1015_Daily_Recovery_Performance_Oct26.zip.gpg"]}
            self.assertIsNone(find_archive(scan,match,"A10XX_Daily_Recovery_Performance_Oct26.zip.gpg"))
            match.update(received_date="261029",mail_date="28102026")
            self.assertEqual(find_archive(scan,match,"A10XX_Daily_Recovery_Performance_Oct26.zip.gpg"),path)

    def test_typographical_keywords_and_decorated_subjects_work_for_every_task(self):
        cases={
            'customer_documents':'FW: [TEST] เอกสารลูกคา รอบ 2 29102026 (เพิ่มเติม)',
            'payment_online':'RE: Online Paymant Pre WO & Post WO with 29/10/2026 Round 2 (TEST)',
            'performance':"FW: Report Daily Recovery Performnce Oct'26 as of 28/10/2026 (test)",
            'performance_success_rate':"[TEST] Performance Succes Rate Oct'26 as of 28/10/2026 เพิ่มเติม",
            'enforcement':'FW: แจ้งผลขออนุมัติยึดอสังหาริมทรัพย_29102026 (TEST)'}
        for task,subject in cases.items():
            with self.subTest(task=task):
                found=match_mail(task,subject,'',received_date='261029')
                self.assertIsNotNone(found)
                self.assertNotIn('trigger_error',found,found)

    def test_body_filename_beats_a_copied_subject_for_another_task(self):
        body='A10XXL_Document_Round2_29102026'
        subject='261029_A1015_A1015_Daily_Recovery_Performance_Oct26.zip.gpg'
        self.assertIsNone(match_mail('performance',subject,body,received_date='261029'))
        self.assertIsNotNone(match_mail('customer_documents',subject,body,received_date='261029'))

    def test_adjust_wrong_office_recall_and_unrelated_mail_are_not_work(self):
        self.assertIsNone(match_mail('performance',"Report Daily Recovery Performance Oct'26 as of 28/10/2026",'Adjust Recovery Performance'))
        self.assertIsNone(match_mail('customer_documents','เอกสารลูกค้า รอบ 2 29102026','A1099L_Document_Round2_29102026.zip.gpg'))
        self.assertIsNone(match_mail('customer_documents','Recall: เอกสารลูกค้า','A10XXL_Document_Round2_29102026'))
        for task in ['payment_online','performance','customer_documents','enforcement','performance_success_rate']:
            self.assertIsNone(match_mail(task,'นัดประชุมทั่วไป','เอกสารสำหรับประชุม'))

    def test_ambiguous_metadata_is_visible_and_never_guessed(self):
        result=match_mail('payment_online','Online Payment Round 2 รอบ 12 29102026','')
        self.assertIn('trigger_error',result)
        result=match_mail('customer_documents','เอกสารลูกค้า','')
        self.assertIn('trigger_error',result)

    def test_round2_never_selects_round12_and_missing_extension_can_find_duplicated_office(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp)
            first=folder/'261029_A1015L_A1015L_Document_Round2_29102026.zip.gpg';first.touch()
            (folder/'261029_A1015L_A1015L_Document_Round12_29102026.zip.gpg').touch()
            self.assertEqual(locate_body_file(folder,'A10XXL_Document_Round2_29102026'),first)
            self.assertEqual(infer_sources(folder,{'round':'2','mail_date':'29102026'},'customer_documents'),[first])
            (folder/'261030_A1015L_A1015L_Document_Round2_29102026.zip.gpg').touch()
            with self.assertRaises(ValidationError):locate_body_file(folder,'A10XXL_Document_Round2_29102026')


class RealScannerFlowTests(unittest.TestCase):
    def test_read_mail_wrong_subject_xx_no_extension_reaches_zip_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp);values=default_settings()
            values.update(logs_root=str(folder/'Logs'),output_root=str(folder/'Completed'))
            values['workflows']['customer_documents']['network_folder']=str(folder)
            values['archive']['archive_processed_files']=False
            values['secrets']={'zip_password':'fixture-zip','passphrase':''}
            values['mail']['workflow_recipients']['customer_documents']={'to':['fixture-documents'],'cc':[]}
            source=folder/'261029_A1015L_A1015L_Document_Round2_29102026.zip.gpg';source.write_bytes(b'encrypted fixture')
            date=now();day=date.strftime('%d%m%Y');stamp=date.strftime('%y%m%d')
            source.rename(folder/(stamp+'_A1015L_A1015L_Document_Round2_'+day+'.zip.gpg'))
            mail=SimpleNamespace(Class=43,EntryID='fixture-read',Subject='เอกสารลูกค้าจาก E-Mail กลาง รอบ2 PM_2902026',
                Body='A10xxL_Document_Round2_'+day,UnRead=False,ReceivedTime=date,Attachments=SimpleNamespace(Count=0))
            items=SimpleNamespace(Count=1,Sort=lambda *a:None,Item=lambda _:mail)
            inbox=SimpleNamespace(Items=items);inbox.Folders=SimpleNamespace(Item=lambda _:inbox)
            store=SimpleNamespace(DisplayName='fixture',StoreID='store',GetRootFolder=lambda:inbox,GetDefaultFolder=lambda _:inbox)
            sent=[];test=self
            class Adapter(Outlook):
                def __init__(self):self.session=SimpleNamespace(Stores=[store])
                def prepare_forward(self,reference,attachment,settings):
                    test.assertEqual(settings['to'],['fixture-documents'])
                    with pyzipper.AESZipFile(attachment) as archive:
                        archive.setpassword(b'fixture-zip');test.assertEqual(archive.read('nested/document.pdf'),b'fixture pdf')
                    return SimpleNamespace(Close=lambda _:None)
                def deliver(self,*args):sent.append('sent');return 'sent'
                def complete(self,*args):pass
            values['workflows']['customer_documents']['store_name']='fixture'
            adapter=Adapter();events=[]
            service=FileWorkflowService(values,'customer_documents',notify=events.append,outlook=adapter)
            def decrypt(command,**kwargs):
                output=Path(command[command.index('--output')+1])
                with zipfile.ZipFile(output,'w') as archive:archive.writestr('nested/document.pdf',b'fixture pdf')
                return SimpleNamespace(returncode=0)
            with patch('routine.files.subprocess.run',side_effect=decrypt):
                result=service.poll()
            self.assertEqual(len(result),1,result)
            if result[0]['status']!='success':self.fail((Path(result[0]['evidence'])/'technical.log').read_text(encoding='utf-8'))
            self.assertEqual(sent,['sent'])
            self.assertEqual(service.poll(),[])
            self.assertEqual(sent,['sent'])

    def test_missing_source_reports_archived_instead_of_replaying(self):
        from routine.intake_status import missing_source_note
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp);archive=folder/'03 Check';archive.mkdir()
            (archive/'261029_A1015L_A1015L_Document_Round1_29102026.zip.gpg').touch()
            result=missing_source_note({'scan':{'network_folder':temp},'archive':{}},
                {'file_names':['A1015L_Document_Round1_29102026']},'customer_documents')
            self.assertIn('03 Check',result)
            self.assertIn('ไม่ส่งซ้ำ',result)

    def test_unready_enforcement_does_not_block_ready_customer_documents(self):
        values=default_settings()
        for task,flow in values['workflows'].items():flow['enabled']=task in ('customer_documents','enforcement')
        values['mail']['workflow_recipients']['customer_documents']={'to':['fixture-documents'],'cc':[]}
        calls=[];events=[]
        fake=SimpleNamespace(settings={'logs':str(Path(tempfile.gettempdir())/'fixture-independent')},poll=lambda:calls.append('polled'))
        with patch('routine.workflows.dependency_checks',side_effect=lambda v,t,**k:['missing Enforcement To'] if t=='enforcement' else []), \
             patch('routine.workflows.FileWorkflowService',return_value=fake),patch('routine.service.Ledger'):
            run(values,threading.Event(),events.append,once=True)
        self.assertEqual(calls,['polled'])
        self.assertTrue(any(e.get('task')=='enforcement' and e.get('phase')=='error' for e in events))


if __name__=='__main__':unittest.main()
