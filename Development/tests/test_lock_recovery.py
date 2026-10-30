"""Real OS handle exclusion, crash release and explicit legacy-marker recovery."""
from pathlib import Path
import json,os,subprocess,sys,tempfile,unittest
from routine.locking import lease,recover_legacy_lock
from routine.vendor.fkp.persistence import read
from routine.vendor.fkp.domain import ValidationError


class LockTests(unittest.TestCase):
    def test_live_lease_cannot_be_acquired_or_recovered_and_cleans_on_exit(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'workflow.lock'
            with lease(temp,'workflow'):
                self.assertEqual(read(path)['version'],2)
                with self.assertRaisesRegex(ValidationError,'งานติดล็อก'):
                    with lease(temp,'workflow'):self.fail('overlapping lease')
                with self.assertRaises(ValidationError):recover_legacy_lock(temp,'workflow',confirmed=True)
                self.assertTrue(path.exists())
            self.assertFalse(path.exists())
            with lease(temp,'workflow'):pass

    def test_process_crash_releases_lease_without_stale_file(self):
        with tempfile.TemporaryDirectory() as temp:
            dev=Path(__file__).resolve().parents[1]
            code=("import sys,os;sys.path.insert(0,sys.argv[1]);from routine.runtime import bootstrap_windows;bootstrap_windows();"
                  "from routine.locking import lease\nwith lease(sys.argv[2],'workflow'):os._exit(0)")
            result=subprocess.run([sys.executable,'-X','utf8','-c',code,str(dev),temp],
                capture_output=True,timeout=20,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            self.assertEqual(result.returncode,0,result.stderr.decode('utf-8','replace'))
            self.assertFalse((Path(temp)/'workflow.lock').exists())
            with lease(temp,'workflow'):pass

    def test_legacy_marker_requires_explicit_recovery_and_preserves_backup(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'workflow.lock';owner={'pid':12345,'started':'2026-10-29T18:22:06+07:00'}
            path.write_text(json.dumps(owner),encoding='utf-8')
            with self.assertRaises(ValidationError):
                with lease(temp,'workflow'):self.fail('legacy marker ignored')
            with self.assertRaises(ValidationError):recover_legacy_lock(temp,'workflow')
            self.assertTrue(path.exists())
            backup=recover_legacy_lock(temp,'workflow',confirmed=True)
            self.assertEqual(read(backup)['owner'],owner)
            self.assertFalse(path.exists())
            with lease(temp,'workflow'):pass

    def test_corrupt_or_new_format_marker_is_never_removed_by_legacy_recovery(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'workflow.lock'
            for content in ['broken',json.dumps({'version':2,'host':'elsewhere','pid':123,'started':'x'})]:
                path.write_text(content,encoding='utf-8')
                with self.assertRaises(ValidationError):recover_legacy_lock(temp,'workflow',confirmed=True)
                self.assertEqual(path.read_text(encoding='utf-8'),content)

    def test_lease_releases_when_body_raises(self):
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaisesRegex(RuntimeError,'fixture'):
                with lease(temp,'workflow'):raise RuntimeError('fixture')
            self.assertFalse((Path(temp)/'workflow.lock').exists())


if __name__=='__main__':unittest.main()
