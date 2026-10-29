"""A saved managed tool path must not depend on a retired release directory."""
from pathlib import Path
import tempfile
import unittest
from routine.settings import default_settings,load,save


class RuntimeUpgradeTests(unittest.TestCase):
    def test_old_bundled_tools_rebase_and_private_settings_survive(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            values=default_settings()
            defaults={k:values[k] for k in ('gpg_path','ocr_path')}
            values['gpg_path']=str(root/'Program/FKP_Routine_System_0.2.1/_internal/runtime/gpg/gpg2.exe')
            values['ocr_path']=str(root/'Program/FKP_Routine_System_0.2.1/_internal/runtime/ocr/tesseract.exe')
            values['gpg_home']=str(root/'private-keys')
            values['secrets']={'passphrase':'fixture-pass','zip_password':'fixture-zip'}
            values['mail']['workflow_recipients']['enforcement']={'to':['fixture-team'],'cc':[]}
            path=root/'profile.json'
            save(values,path)
            result=load(path)
            self.assertEqual({k:result[k] for k in defaults},defaults)
            self.assertEqual(result['gpg_home'],values['gpg_home'])
            self.assertEqual(result['secrets'],values['secrets'])
            self.assertEqual(result['mail']['workflow_recipients']['enforcement']['to'],['fixture-team'])

    def test_legacy_mail_window_is_migrated_into_visible_scan_controls(self):
        import json
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);values=default_settings()
            path=root/"profile.json";save(values,path)
            data=json.loads(path.read_text(encoding="utf-8"))
            data.pop("scan_hours");data.pop("scan_limit")
            data["outlook"].update(received_within_hours=168,max_scan_items=600)
            path.write_text(json.dumps(data),encoding="utf-8")
            result=load(path)
            self.assertEqual((result["scan_hours"],result["scan_limit"]),(168,600))

    def test_custom_tool_paths_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            values=default_settings()
            values['gpg_path']=str(root/'Custom/GnuPG/gpg.exe')
            values['ocr_path']=str(root/'Custom/Tesseract/tesseract.exe')
            path=root/'profile.json';save(values,path)
            loaded=load(path)
            for key in ('gpg_path','ocr_path'):
                self.assertEqual(loaded[key],values[key])


if __name__=='__main__':unittest.main()
