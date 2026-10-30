"""Document failures remain visible and are recorded before a job exists."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import hashlib,tempfile,threading,unittest
from routine.settings import default_settings,task_config,runtime_settings
from routine.monitoring import MonitorState,record_scan_issue
from routine.intake_status import missing_source_note
from routine.outlook import MailRef
from routine.service import mail_key
from routine.workflows import FileWorkflowService,run
from routine.vendor.fkp.domain import ValidationError
from routine.vendor.fkp.persistence import now,read,digest


class DocumentStatusTests(unittest.TestCase):
    def values(self,root):
        values=default_settings();values['logs_root']=str(root/'Logs');values['output_root']=str(root/'Completed')
        for task,flow in values['workflows'].items():flow['enabled']=task=='customer_documents'
        values['workflows']['customer_documents']['network_folder']=str(root)
        return values

    def test_waiting_file_name_stays_visible_during_next_poll_countdown(self):
        model=MonitorState(default_settings()['workflows'])
        model.apply({'kind':'status','task':'customer_documents','attention':True,
                     'step':'รอไฟล์ A1015L_Document_Round1_30102026.zip.gpg'})
        model.apply({'kind':'poll_wait','task':'customer_documents','next_check':now().isoformat()})
        self.assertIn('A1015L_Document_Round1_30102026',model.visible_step('customer_documents'))
        self.assertEqual(model.rows['customer_documents']['phase'],'attention')
        model.stopped()
        self.assertIn('A1015L_Document_Round1_30102026',model.visible_step('customer_documents'))

    def test_missing_source_note_names_the_file_and_archive_does_not_replay_it(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);name='A1015L_Document_Round1_30102026.zip.gpg'
            runtime={'scan':{'network_folder':str(root)},'archive':{}}
            self.assertIn(name,missing_source_note(runtime,{},'customer_documents',name))
            archive=root/'03 Check';archive.mkdir();(archive/name).write_bytes(b'completed elsewhere')
            note=missing_source_note(runtime,{},'customer_documents',name)
            self.assertIn('03 Check',note);self.assertIn(name,note)
            self.assertFalse((root/name).exists())

    def test_failed_previous_job_explains_stage_without_retrying_or_sending(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);values=self.values(root);name='A1015L_Document_Round1_30102026.zip.gpg'
            source=root/name;source.write_bytes(b'fixture')
            ref=MailRef('entry','store','เอกสารลูกค้า','',{'file_names':[name]})
            events=[]
            outlook=SimpleNamespace(scan_workflow=lambda _:[ref],deliver=lambda *a:self.fail('mail must not be sent'))
            service=FileWorkflowService(values,'customer_documents',notify=events.append,outlook=outlook)
            fingerprint=digest(source);key=hashlib.sha256((mail_key(ref)+'\0'+name.casefold()).encode()).hexdigest()
            service.ledger.data['files'][fingerprint]={'status':'failed','mail':'not_requested','last_error':'สร้าง ZIP: fixture error','job':'fixture/job'}
            service.ledger.data['mails'][key]={'file':fingerprint,'status':'failed'};service.ledger.save()
            with patch.object(service,'process_unit',side_effect=AssertionError('must not retry')):
                service.poll()
            attention=[event for event in events if event.get('attention')]
            self.assertEqual(len(attention),1)
            self.assertIn('สร้าง ZIP',attention[0]['step'])
            self.assertIn('fixture/job',attention[0]['step'])

    def test_run_error_before_job_is_saved_for_later_diagnosis(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);values=self.values(root);events=[]
            with patch('routine.workflows.dependency_checks',return_value=[]),\
                 patch.object(FileWorkflowService,'poll',side_effect=ValidationError('งานติดล็อก fixture')):
                run(values,threading.Event(),events.append,once=True)
            path=Path(task_config(values,'customer_documents')['logs'])/'Logs/last_scan_issue.json'
            issue=read(path)
            self.assertEqual(issue['task'],'customer_documents')
            self.assertEqual(issue['phase'],'error')
            self.assertIn('งานติดล็อก',issue['step'])


if __name__=='__main__':unittest.main()
