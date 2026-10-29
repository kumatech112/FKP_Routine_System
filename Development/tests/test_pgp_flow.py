"""PGP automatic key-store use must not be replaced by an empty password."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import tempfile
import unittest
import zipfile

from routine.files import extract_payload
from routine.settings import default_settings, validate
from routine.vendor.fkp.domain import ValidationError


class PgpAutomaticFlowTests(unittest.TestCase):
    def exercise(self, extension, homedir=""):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / ("payload.zip." + extension)
            source.write_bytes(b"encrypted fixture")
            job = SimpleNamespace(folder=root / "Process", check=lambda: None)
            job.folder.mkdir()
            commands = []
            def run(command, **kwargs):
                commands.append(command)
                self.assertNotIn("--dump-options", command)
                self.assertNotIn("--passphrase-fd", command)
                self.assertNotIn("--pinentry-mode", command)
                self.assertNotIn("input", kwargs)
                self.assertEqual("--homedir" in command, bool(homedir))
                if homedir:
                    self.assertEqual(command[command.index("--homedir")+1], homedir)
                output = Path(command[command.index("--output")+1])
                with zipfile.ZipFile(output, "w") as archive:
                    archive.writestr("result.txt", b"decrypted content")
                return SimpleNamespace(returncode=0)
            with patch("routine.files.subprocess.run", side_effect=run):
                folder = extract_payload(source, job, {"enabled":True, "gpg_path":"fixture.exe",
                    "passphrase":"", "gpg_home":homedir})
            self.assertEqual((folder/"result.txt").read_bytes(), b"decrypted content")
            self.assertEqual(len(commands), 1)
            self.assertTrue(source.exists())

    def test_pgp_uses_machine_key_store_without_empty_password_override(self):
        self.exercise("pgp")

    def test_gpg_accepts_explicit_key_store_without_password_override(self):
        self.exercise("gpg", r"C:\Fixture\keys")

    def test_failed_key_store_decryption_never_extracts_payload(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            source=root/"payload.zip.pgp"
            source.write_bytes(b"encrypted fixture")
            job=SimpleNamespace(folder=root/"Process", check=lambda: None)
            job.folder.mkdir()
            with patch("routine.files.subprocess.run", return_value=SimpleNamespace(returncode=2)), \
                 patch("routine.files.extract_zip") as extract:
                with self.assertRaises(ValidationError):
                    extract_payload(source,job,{"enabled":True,"gpg_path":"fixture.exe"})
                extract.assert_not_called()
            self.assertTrue(source.exists())

    def test_machine_key_store_and_empty_password_are_valid_settings(self):
        values=default_settings()
        self.assertEqual(values["gpg_home"], "")
        self.assertEqual(validate(values)["secrets"]["passphrase"], "")

    def test_explicit_key_store_still_requires_absolute_path(self):
        values=default_settings()
        values["gpg_home"]="relative-key-folder"
        with self.assertRaises(ValidationError):
            validate(values)


if __name__=="__main__":
    unittest.main()
