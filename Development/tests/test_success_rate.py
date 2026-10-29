"""Body-driven Success Rate intake and real RAR-to-password-ZIP delivery, isolated mail."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import tempfile
import json
import unittest
import zipfile
import pyzipper

from routine.files import extract_payload
from routine.matching import success_rate_match, success_rate_file
from routine.outlook import MailRef
from routine.settings import default_settings,runtime_settings,save,load
from routine.workflows import FileWorkflowService
from routine.vendor.fkp.domain import ValidationError

TITLE="Performance success rate ปิดหนี้ไวไปต่อได้ Oct'26 สำหรับสำนักงานกฎหมาย as of 28/10/2026"
BODY="File Name: A10XX_Performance_SAM_OA_Oct26.zip.gpg"
FIXTURES=Path(__file__).parent/'fixtures'


class MatchingTests(unittest.TestCase):
    def test_body_placeholder_selects_office15_received_date_not_previous_monthly_delivery(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp)
            names=['261028_A1015_A1015_Performance_SAM_OA_Oct26.zip.gpg',
                   '261029_A1015_A1015_Performance_SAM_OA_Oct26.zip.gpg',
                   '261029_A1099_A1099_Performance_SAM_OA_Oct26.zip.gpg']
            for name in names:(folder/name).write_bytes(b'fixture')
            match=success_rate_match('[TEST] '+TITLE+' (additional)',BODY)
            match['received_date']='261029'
            self.assertEqual(success_rate_file({'network_folder':temp},match,BODY).name,names[1])

    def test_rar_name_and_a1015l_are_supported(self):
        body=BODY.replace('A10XX','A1015L').replace('.zip.','.rar.')
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/body.split(': ')[1];path.write_bytes(b'fixture')
            self.assertEqual(success_rate_file({'network_folder':temp},success_rate_match(TITLE,body),body),path)

    def test_wrong_month_or_ambiguous_files_never_select_arbitrarily(self):
        with self.assertRaises(ValidationError):success_rate_match(TITLE,BODY.replace('Oct26','Sep26'))
        with tempfile.TemporaryDirectory() as temp:
            for ext in ['gpg','pgp']:(Path(temp)/('A1015_Performance_SAM_OA_Oct26.zip.'+ext)).touch()
            with self.assertRaises(ValidationError):success_rate_file({'network_folder':temp},success_rate_match(TITLE,BODY),BODY)

    def test_old_four_task_profile_adds_disabled_task_without_recipient_fallback(self):
        with tempfile.TemporaryDirectory() as temp:
            values=default_settings()
            values['secrets']['zip_password']='fixture-password'
            path=Path(temp)/'profile.json';save(values,path)
            old=json.loads(path.read_text(encoding='utf-8'))
            old['workflows'].pop('performance_success_rate')
            old['mail']['workflow_recipients'].pop('performance_success_rate')
            path.write_text(json.dumps(old),encoding='utf-8')
            actual=load(path)
            self.assertFalse(actual['workflows']['performance_success_rate']['enabled'])
            self.assertEqual(actual['mail']['workflow_recipients']['performance_success_rate'],{'to':[],'cc':[]})
            self.assertEqual(actual['secrets'],values['secrets'])
            for task in values['workflows']:self.assertEqual(actual['workflows'][task],values['workflows'][task])


class PayloadTests(unittest.TestCase):
    def test_pgp_payload_detects_rar4_and_rar5_by_content_and_preserves_nested_unicode_bytes(self):
        for version in [4,5]:
            with self.subTest(version=version), tempfile.TemporaryDirectory() as temp:
                root=Path(temp);source=root/'payload.zip.pgp';source.write_bytes(b'encrypted fixture')
                work=root/'Process';work.mkdir()
                def decrypt(command,**kwargs):
                    Path(command[command.index('--output')+1]).write_bytes((FIXTURES/f'payload_rar{version}.rar').read_bytes())
                    return SimpleNamespace(returncode=0)
                with patch('routine.files.subprocess.run',side_effect=decrypt):
                    # Only GPG is mocked; the actual UnRAR subprocess is launched by rarfile.Popen.
                    folder=extract_payload(source,SimpleNamespace(folder=work,check=lambda:None),{'enabled':True,'gpg_path':'fixture.exe'})
                self.assertEqual((folder/'nested/result.txt').read_bytes(),b'fixture content '*1000)
                self.assertEqual((folder/'รายงาน.txt').read_text(encoding='utf-8'),'ข้อมูลทดสอบ RAR')
                self.assertTrue(source.exists())

    def test_unsafe_rar_metadata_is_rejected_before_opening_any_member(self):
        from routine.rar_extract import extract_rar
        from unittest.mock import MagicMock
        for names in [['../outside.txt'],['C:/outside.txt'],['same.txt','SAME.txt'],['link.txt']]:
            with self.subTest(names=names),tempfile.TemporaryDirectory() as temp:
                archive=MagicMock();archive.__enter__.return_value=archive
                archive.needs_password.return_value=False;archive.volumelist.return_value=['one.rar']
                archive.infolist.return_value=[SimpleNamespace(filename=n,file_size=10,
                    is_symlink=lambda n=n:n=='link.txt',is_dir=lambda:False,file_redir=None) for n in names]
                with patch('routine.rar_extract.rarfile.RarFile',return_value=archive):
                    with self.assertRaises(ValidationError):extract_rar('fixture.rar',Path(temp)/'Extracted')
                archive.open.assert_not_called()
                self.assertEqual(list(Path(temp).iterdir()),[])

    def test_corrupt_rar_payload_fails_crc_validation(self):
        import rarfile
        from routine.rar_extract import extract_rar
        with tempfile.TemporaryDirectory() as temp:
            source=Path(temp)/'corrupt.rar';data=bytearray((FIXTURES/'payload_rar4.rar').read_bytes())
            with rarfile.RarFile(FIXTURES/'payload_rar4.rar') as archive:
                info=next(i for i in archive.infolist() if not i.is_dir())
                offset=info.data_offset
            data[offset+2]^=0xff;source.write_bytes(data)
            with self.assertRaises(ValidationError):extract_rar(source,Path(temp)/'Extracted')

    def test_damaged_archive_stops_before_any_delivery(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);source=root/'broken.rar'
            source.write_bytes((FIXTURES/'payload_rar5.rar').read_bytes()[:80])
            work=root/'Process';work.mkdir()
            with self.assertRaises(ValidationError):extract_payload(source,SimpleNamespace(folder=work,check=lambda:None),{})

    def test_success_rate_rar_is_repacked_encrypted_and_uses_own_recipients_no_database(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);values=default_settings()
            values.update(logs_root=str(root/'Logs'),output_root=str(root/'Completed'))
            values['archive']['archive_processed_files']=False
            values['secrets']['zip_password']='fixture-zip'
            values['mail']['workflow_recipients']['performance_success_rate']={'to':['success-team'],'cc':['success-cc']}
            values['mail']['send_immediately']=True
            source=root/'A1015_Performance_SAM_OA_Oct26.rar';source.write_bytes((FIXTURES/'payload_rar5.rar').read_bytes())
            calls=[]
            def prepare(ref,attachment,mail):
                self.assertEqual(mail['to'],['success-team']);self.assertEqual(mail['cc'],['success-cc'])
                with pyzipper.AESZipFile(attachment) as archive:
                    with self.assertRaises(RuntimeError):archive.read('nested/result.txt')
                    archive.setpassword(b'fixture-zip')
                    self.assertEqual(archive.read('nested/result.txt'),b'fixture content '*1000)
                    self.assertEqual(set(archive.namelist()),{'nested/result.txt','รายงาน.txt'})
                calls.append('prepared');return SimpleNamespace(Close=lambda _:None)
            def deliver(*args):calls.append('sent');return 'sent'
            outlook=SimpleNamespace(prepare_forward=prepare,deliver=deliver,complete=lambda *a:None)
            service=FileWorkflowService(values,'performance_success_rate',outlook=outlook)
            ref=MailRef('entry','store',TITLE,BODY,{})
            result=service.process_unit(ref,source,runtime_settings(values,'performance_success_rate'),outlook,'unit')
            self.assertEqual(result['status'],'success',result)
            self.assertEqual(calls,['prepared','sent'])
            state=next(iter(service.ledger.data['files'].values()))
            self.assertEqual(state['database'],'not_requested')
            self.assertIsNone(service.process_unit(ref,source,runtime_settings(values,'performance_success_rate'),outlook,'repeat'))
            self.assertEqual(calls.count('sent'),1)


if __name__=='__main__':unittest.main()
