"""Observe real scan transitions and fail closed when Outlook changes the selected sender."""
from datetime import timedelta
from types import SimpleNamespace
from unittest.mock import patch
import unittest
from routine.monitoring import MonitorState,scan_mail
from routine.outlook import Outlook,require_sender,list_sender_accounts
from routine.settings import default_settings
from routine.vendor.fkp.domain import ValidationError
from routine.vendor.fkp.persistence import now


class MonitorTests(unittest.TestCase):
    def test_zero_matches_only_appears_after_actual_scan_and_wait_countdown_is_timestamp_based(self):
        model=MonitorState(default_settings()['workflows']);events=[]
        self.assertIsNone(model.rows['enforcement']['matched'])
        settings={'store_name':'fixture','folder_path':['Inbox']}
        scan_mail(SimpleNamespace(scan_workflow=lambda _:[]),settings,events.append,'enforcement')
        self.assertEqual([e['phase'] for e in events],['checking','checked'])
        for event in events:model.apply(event)
        self.assertEqual(model.rows['enforcement']['matched'],0)
        stamp=now();model.apply({'kind':'poll_wait','task':'enforcement','next_check':(stamp+timedelta(seconds=25)).isoformat()})
        self.assertIn('25',model.visible_step('enforcement',stamp))
        self.assertEqual(model.visible_step('enforcement',stamp+timedelta(seconds=26)),'รอคิวตรวจเมล')
        model.stopped();self.assertEqual(model.rows['enforcement']['phase'],'stopped')
        self.assertEqual(model.rows['performance_success_rate']['phase'],'disabled')

    def test_scan_error_does_not_claim_zero_matches_or_successful_poll(self):
        events=[];model=MonitorState(default_settings()['workflows'])
        def fail(_):raise RuntimeError('COM unavailable')
        with self.assertRaises(ValidationError):scan_mail(SimpleNamespace(scan_workflow=fail),
            {'store_name':'fixture','folder_path':['Inbox']},events.append,'enforcement')
        for event in events:model.apply(event)
        model.apply({'kind':'poll_wait','task':'enforcement','next_check':now().isoformat()})
        self.assertEqual(model.rows['enforcement']['phase'],'error')
        self.assertIsNone(model.rows['enforcement']['last_checked'])
        self.assertIsNone(model.rows['enforcement']['matched'])


class SenderTests(unittest.TestCase):
    def test_explicit_exact_smtp_is_required_and_duplicates_are_rejected(self):
        account=SimpleNamespace(SmtpAddress='sender@example.test')
        mail=SimpleNamespace(SendUsingAccount=None)
        require_sender(mail,[account],' SENDER@example.test ')
        self.assertIs(mail.SendUsingAccount,account)
        for accounts,sender in [([account],'missing@example.test'),([account,account],account.SmtpAddress),([account],'')]:
            with self.assertRaises(ValidationError):require_sender(mail,accounts,sender)

    def test_assignment_rejected_by_outlook_blocks_send(self):
        class IgnoringMail:
            @property
            def SendUsingAccount(self):return SimpleNamespace(SmtpAddress='wrong@example.test')
            @SendUsingAccount.setter
            def SendUsingAccount(self,_):pass
        with self.assertRaises(ValidationError):require_sender(IgnoringMail(),[SimpleNamespace(SmtpAddress='sender@example.test')],'sender@example.test')

    def test_draft_display_and_save_must_keep_chosen_account(self):
        chosen=SimpleNamespace(SmtpAddress='sender@example.test')
        adapter=Outlook.__new__(Outlook);adapter.session=SimpleNamespace(Accounts=[chosen])
        for operation in ['Display','Save']:
            calls=[];mail=SimpleNamespace(SendUsingAccount=None)
            def mutate(*args):mail.SendUsingAccount=SimpleNamespace(SmtpAddress='wrong@example.test')
            mail.Display=mutate if operation=='Display' else lambda *a:None
            mail.Save=mutate if operation=='Save' else lambda:None
            mail.Close=lambda _:calls.append('closed')
            mail.Send=lambda:calls.append('sent')
            with self.assertRaises(ValidationError):adapter.deliver(mail,{'send_from':chosen.SmtpAddress,'send_immediately':False})
            self.assertEqual(calls,['closed'])

    def test_account_listing_is_explicit_sorted_and_ignores_blank(self):
        import win32com.client
        accounts=[SimpleNamespace(SmtpAddress=s) for s in ['b@example.test','','a@example.test','b@example.test']]
        with patch.object(win32com.client,'Dispatch',return_value=SimpleNamespace(Session=SimpleNamespace(Accounts=accounts))):
            self.assertEqual(list_sender_accounts(),['a@example.test','b@example.test'])


if __name__=='__main__':unittest.main()
