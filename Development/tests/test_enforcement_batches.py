"""Regression coverage for same-day multi-file intake and sequential batch preparation."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import json,tempfile,unittest
import test_enforcement_delivery as fixtures
from routine.triggers import infer_sources,match_mail
from routine.workflows import FileWorkflowService,run
from routine.outlook import MailRef
from routine.settings import default_settings,runtime_settings,dependency_checks,save,load
from routine.vendor.fkp.domain import ValidationError

class InferenceTests(unittest.TestCase):
    def test_same_day_four_numbered_sources_are_one_candidate_batch(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp)
            names=[f"261029_A1015L_A1015L_Approval_Enforcement_BAY_29102026{suffix}.pdf.gpg" for suffix in ("","_2","_3","_4")]
            for name in names:(folder/name).touch()
            for other in ("A1099L_Approval_Enforcement_BAY_29102026.pdf.gpg",
                "A1015L_Reject_Enforcement_29102026.pdf.gpg","A1015L_Approval_Enforcement_BAY_28102026.pdf.gpg"):
                (folder/other).touch()
            match=match_mail("enforcement","แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์_29102026","")
            self.assertEqual({p.name for p in infer_sources(folder,match,"enforcement")},set(names))
    def test_partial_inline_ocr_names_do_not_drop_other_same_day_sources(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp);values=default_settings();values.update(logs_root=str(folder/'Logs'),output_root=str(folder/'Completed'))
            values['workflows']['enforcement']['network_folder']=str(folder)
            files=[]
            for suffix in ('','_2','_3','_4'):
                path=folder/('261029_A1015L_Approval_Enforcement_BAY_29102026'+suffix+'.pdf.gpg');path.touch();files.append(path)
            ref=MailRef('inline','store','แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์_29102026','',match_mail('enforcement','แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์_29102026',''))
            adapter=SimpleNamespace(inline_images=lambda *a:[folder/'fixture.png'])
            service=FileWorkflowService(values,'enforcement',outlook=adapter)
            with patch('routine.workflows.image_mail_text',return_value=files[0].name+' '+files[2].name):
                self.assertEqual({path for path,name in service.candidates(ref,runtime_settings(values,'enforcement'),adapter)},set(files))
            explicit=MailRef('explicit','store',ref.subject,files[0].name,{'file_names':[files[0].name]})
            adapter.inline_images=lambda *a:[]
            self.assertEqual([path for path,name in service.candidates(explicit,runtime_settings(values,'enforcement'),adapter)],[files[0]])
    def test_one_source_and_equivalent_plain_copy_choose_encrypted(self):
        with tempfile.TemporaryDirectory() as temp:
            folder=Path(temp);base="261029_A1015L_Approval_Enforcement_29102026.pdf"
            (folder/base).touch();encrypted=folder/(base+".gpg");encrypted.touch()
            self.assertEqual(infer_sources(folder,{"decision":"Approval","mail_date":"29102026"},"enforcement"),[encrypted])
            (folder/(base+".pgp")).touch()
            with self.assertRaises(ValidationError):infer_sources(folder,{"decision":"Approval","mail_date":"29102026"},"enforcement")
    def test_duplicate_transport_dates_remain_ambiguous(self):
        with tempfile.TemporaryDirectory() as temp:
            for stamp in ("261029","261030"):
                (Path(temp)/(stamp+"_A1015L_Approval_Enforcement_29102026.pdf.gpg")).touch()
            with self.assertRaises(ValidationError):infer_sources(temp,{"decision":"Approval","mail_date":"29102026"},"enforcement")

class SequentialTests(unittest.TestCase):
    def setUp(self):
        self.fixture=fixtures.DeliveryTests("runTest");self.fixture.setUp()
    def tearDown(self):self.fixture.tearDown()
    def test_actual_poll_prepares_four_sets_in_order_and_sends_one_flat_zip(self):
        f=self.fixture;f.expected_excel_count=4
        f.values["workflows"]["enforcement"]["network_folder"]=str(f.root)
        # The initial setup source uses another date; it must not enter this batch.
        for suffix in ("","_2","_3","_4"):
            (f.root/("261029_A1015L_A1015L_Approval_Enforcement_BAY_29102026"+suffix+".pdf.gpg")).write_bytes(("encrypted"+suffix).encode())
        ref=MailRef("four","store","แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์_29102026","",
                    match_mail("enforcement","แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์_29102026",""))
        f.outlook.scan_workflow=lambda _:[ref];f.outlook.inline_images=lambda *a:[]
        f.service=FileWorkflowService(f.values,"enforcement",outlook=f.outlook)
        with patch("routine.enforcement_flow.extract_payload",side_effect=f.extract),patch("routine.enforcement_flow.prepare_enforcement",side_effect=f.ocr):
            result=f.service.poll()
            self.assertEqual(result[0]["status"],"success",result)
            self.assertEqual(result[0]["result"]["sets"],4)
            self.assertEqual(len(result[0]["result"]["outputs"]),5)
            self.assertEqual(f.service.poll(),[])
        self.assertEqual(f.events,["extract","ocr"]*4+["prepared","mail"])
    def test_failure_in_second_set_never_delivers_first_set(self):
        f=self.fixture;other=f.root/"A1015L_Approval_Enforcement_28102026_2.pdf.gpg";other.write_bytes(b"other")
        calls=[]
        def ocr(*args):
            calls.append(1)
            if len(calls)==2:raise ValidationError("second set failed")
            return f.ocr(*args)
        receipt=f.execute(ocr=ocr,sources=[f.source,other])
        self.assertEqual(receipt["status"],"failed")
        self.assertEqual(f.state()["prepared_sets"],1)
        self.assertNotIn("mail",f.events)
        self.assertTrue(f.source.exists() and other.exists())
    def test_missing_recipient_prepares_all_sets_and_resume_does_not_repeat_ocr(self):
        f=self.fixture;f.expected_excel_count=2
        other=f.root/"A1015L_Approval_Enforcement_28102026_2.pdf.gpg";other.write_bytes(b"other")
        f.values["mail"]["workflow_recipients"]["enforcement"]={"to":[],"cc":[]}
        receipt=f.execute(sources=[f.source,other])
        self.assertEqual(receipt["status"],"failed")
        self.assertIn("OCR ครบแล้ว",receipt["details"])
        self.assertEqual(f.state()["status"],"ready_to_send")
        self.assertEqual(f.state()["prepared_sets"],2)
        self.assertEqual(f.events,["extract","ocr"]*2)
        batch=json.loads(Path(f.state()["review"]).read_text(encoding="utf-8"))
        self.assertTrue(all(Path(item["output"]).is_file() for item in batch["sets"]))
        f.values["mail"]["workflow_recipients"]["enforcement"]={"to":["fixture-enforcement"],"cc":["fixture-review"]}
        self.assertEqual(f.resume()["status"],"success")
        self.assertEqual(f.events,["extract","ocr"]*2+["prepared","mail"])
    def test_duplicate_pdf_and_excel_basenames_are_both_retained(self):
        f=self.fixture;f.expected_excel_count=2
        other=f.source.with_suffix(".gpg");other.write_bytes(b"other")
        receipt=f.execute(sources=[f.source,other])
        self.assertEqual(receipt["status"],"success",receipt)
        outputs=receipt["result"]["outputs"]
        self.assertEqual(len(outputs),3)
        self.assertEqual(len({Path(p).name.casefold() for p in outputs}),3)
    def test_multiple_pdfs_from_one_source_are_one_ocr_set(self):
        f=self.fixture
        def extract(*args,**kwargs):
            folder=f.extract(*args,**kwargs)
            (folder/"two/case.pdf").write_bytes(b"fixture pdf")
            return folder
        service=FileWorkflowService(f.values,"enforcement",outlook=f.outlook)
        with patch("routine.enforcement_flow.extract_payload",side_effect=extract),patch("routine.enforcement_flow.prepare_enforcement",side_effect=f.ocr):
            receipt=service.process_unit(f.reference,f.source,runtime_settings(f.values,"enforcement"),f.outlook,"one-archive")
        self.assertEqual(receipt["status"],"success",receipt)
        self.assertEqual(receipt["result"]["pdfs"],2)
        self.assertEqual(receipt["result"]["sets"],1)
        self.assertEqual(f.events,["extract","ocr","prepared","mail"])

class RecipientPersistenceTests(unittest.TestCase):
    def test_saved_enforcement_recipients_survive_reload(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/"settings.json";values=default_settings()
            values["mail"]["workflow_recipients"]["enforcement"]={"to":["fixture-group"],"cc":["fixture-cc"]}
            save(values,path);again=load(path)
            self.assertEqual(again["mail"]["workflow_recipients"]["enforcement"],{"to":["fixture-group"],"cc":["fixture-cc"]})
            self.assertEqual(runtime_settings(again,"enforcement")["mail"]["to"],["fixture-group"])
    def test_other_incomplete_workflow_cannot_discard_saved_enforcement_recipients(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'settings.json';values=default_settings()
            values['workflows']['performance_success_rate']['enabled']=True
            values['mail']['workflow_recipients']['performance_success_rate']={'to':[],'cc':[]}
            values['mail']['workflow_recipients']['enforcement']={'to':['fixture-group'],'cc':['fixture-cc']}
            save(values,path)
            self.assertEqual(load(path)['mail']['workflow_recipients']['enforcement']['to'],['fixture-group'])
            self.assertTrue(any('ผู้รับ' in text for text in dependency_checks(values,'performance_success_rate')))
    def test_delivery_readiness_can_be_separated_from_preparation(self):
        values=default_settings()
        values["mail"]["workflow_recipients"]["enforcement"]={"to":[],"cc":[]}
        prepare=dependency_checks(values,"enforcement",include_delivery=False)
        deliver=dependency_checks(values,"enforcement",include_delivery=True)
        self.assertFalse(any("Enforcement To" in text for text in prepare))
        self.assertTrue(any("Enforcement To" in text for text in deliver))

if __name__=="__main__":unittest.main()
