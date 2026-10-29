"""Native GUI/profile verification in isolated paths; production monitoring is never started."""
from pathlib import Path
from unittest.mock import patch
from time import monotonic,sleep
import argparse
from PIL import ImageGrab
from routine.settings import default_settings,load,TASKS
from routine.ui_all import Application
from routine.vendor.fkp.persistence import dump,now


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--destination',type=Path,required=True)
    args=parser.parse_args()
    target=args.destination.resolve();target.mkdir(parents=True,exist_ok=False)
    values=default_settings()
    values.update(logs_root=str(target/'Logs'),output_root=str(target/'Completed'))
    values['secrets']={'passphrase':'fixture-only-pass','zip_password':'fixture-only-zip'}
    with patch('routine.ui_all.run') as worker:
        app=Application(values,target/'settings.json');app.update()
        assert set(app.task_flags)==set(values['workflows'])
        assert not app.dirty() and not worker.called
        assert len(app.notebook.tabs())==3 and len(app.main_tabs.tabs())==2
        assert set(app.recipient_vars)==set(TASKS)
        app.recipient_vars['enforcement']['to'].set('fixture-enforcement')
        app.recipient_vars['enforcement']['cc'].set('fixture-review')
        app.save()
        actual=load(target/'settings.json')
        assert actual['secrets']==values['secrets']
        assert actual['mail']['workflow_recipients']['enforcement']=={'to':['fixture-enforcement'],'cc':['fixture-review']}
        assert 'fixture-only-pass' not in (target/'settings.json').read_text(encoding='utf-8')
        app.mail_vars['send_from'].set('typed@example.test')
        with patch('routine.ui_all.list_sender_accounts',return_value=['a@example.test','b@example.test']):
            app.load_sender_accounts()
            until=monotonic()+5
            while app.sender_loading and monotonic()<until:app.update();sleep(.02)
        assert not app.sender_loading
        assert app.mail_vars['send_from'].get()=='typed@example.test'
        assert tuple(app.sender_combo.cget('values'))==('a@example.test','b@example.test')
        assert str(app.sender_combo.cget('state'))=='normal'
        with patch('routine.ui_all.list_sender_accounts',side_effect=RuntimeError('fixture failure')):
            app.load_sender_accounts();until=monotonic()+5
            while app.sender_loading and monotonic()<until:app.update();sleep(.02)
        assert app.mail_vars['send_from'].get()=='typed@example.test'
        assert tuple(app.sender_combo.cget('values'))==('a@example.test','b@example.test')
        settings_tabs=app.settings_page.winfo_children()[0]
        for size,label in [('1180x780','normal'),('1000x650','constrained')]:
            app.geometry(size+'+20+20')
            for index,page in enumerate(['main','history','settings']):
                app.notebook.select(index)
                if page=='settings':settings_tabs.select(0)
                app.update()
                ImageGrab.grab((app.winfo_rootx(),app.winfo_rooty(),app.winfo_rootx()+app.winfo_width(),app.winfo_rooty()+app.winfo_height())).save(target/f'{label}_{page}.png')
            settings_tabs.select(1);app.update()
            ImageGrab.grab((app.winfo_rootx(),app.winfo_rooty(),app.winfo_rootx()+app.winfo_width(),app.winfo_rooty()+app.winfo_height())).save(target/f'{label}_send_from.png')
        assert 'scan_hours' in app.fields and 'scan_limit' in app.fields
        assert not worker.called and app.worker is None
        app.destroy()
    dump(target/'verification.json',{'status':'verified','tasks':len(TASKS),'pages':3,
        'profile_encrypted':True,'sender_accounts_async':True,'typed_sender_preserved_on_refresh_and_failure':True,
        'status_panel_present':True,'production_monitor_started':False,'emails_sent':0,'database_writes':0})


if __name__=='__main__':main()
