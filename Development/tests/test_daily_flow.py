"""Boundary tests: Daily/Adjust rules, stage routing and irreversible outcomes."""
from dataclasses import replace
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import copy
import json
import tempfile
import threading
import unittest
import zipfile

from routine.config import DAILY_PRODUCTS, default_config, engine_settings, load_config
from routine.daily import DailyArtifacts, DailyDatabase, process_daily, records_from_05
from routine.files import extract_zip, find_archive, make_attachment
from routine.outlook import MailRef
from routine.service import PerformanceService, RoutineJob, mail_key
from routine.vendor.fkp.daily import process_plan
from routine.vendor.fkp.domain import Detail, ValidationError
from routine.vendor.fkp.persistence import digest


def record():
    return Detail("daily", "OCT-26", "2026-10-27", "OA 15", 200,
                  Decimal("1000.0000"), Decimal("50.0000"), product="FC",
                  bucket="B01", efficiency=Decimal("5.0000"))


class FakeSql:
    def __init__(self, fail_insert=False, bad_readback=False, initial=()):
        self.rows = copy.deepcopy(list(initial))
        self.fail_insert, self.bad_readback = fail_insert, bad_readback
        self.commits = self.rollbacks = 0
        self.statements = []
        self.snapshot = None

    def run(self, sql, args=(), rows=True):
        import re
        self.statements.append(sql)
        if sql.startswith("SELECT"):
            result = copy.deepcopy(self.rows)
            if self.bad_readback and self.commits and result:
                result[0]["Collected"] = Decimal("999")
            return result
        if sql.startswith("INSERT"):
            if self.fail_insert:
                raise RuntimeError("insert failure")
            fields = re.findall(r"\[([^]]+)\]", sql.split("VALUES")[0])[2:]
            self.rows.append(dict(zip(fields, args)))
            return []
        raise AssertionError(sql)

    def begin(self):
        self.snapshot = copy.deepcopy(self.rows)

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rows = self.snapshot
        self.rollbacks += 1

    def close(self):
        pass


class FlowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.config = default_config()
        self.config.update(logs=str(self.root / "Performance"), output=str(self.root / "Completed"))
        self.settings = engine_settings(self.config)
        self.job = RoutineJob(self.root / "job", self.settings)
        self.source = self.root / "261028_A1015_A1015_Daily_Recovery_Performance_Oct26.zip"
        with zipfile.ZipFile(self.source, "w") as archive:
            archive.writestr("Daily_Recovery_Performance_Oct26.xlsm", b"fixture only")
        self.legacy = {
            "scan": {"excel_regex": r"Daily_Recovery_Performance_Oct26\.xlsm"},
            "decrypt": {"enabled": False},
            "mail": {"zip_attachment": False, "send_immediately": True},
            "archive": {"archive_processed_files": False},
        }
        self.reference = MailRef("entry1", "store1", "Daily", "", {
            "mail_date": "27/10/2026", "month_name": "Oct", "year": "26",
        })

    def tearDown(self):
        self.temp.cleanup()

    def artifacts(self, source, settings, job):
        a, b = job.folder / "01_source_Oct26.xlsx", job.folder / "05_source_Oct26.xlsx"
        a.write_bytes(b"stage01")
        b.write_bytes(b"stage05")
        return DailyArtifacts(a, b, (record(),), digest(a), digest(b))

    def service(self, sql=None, fail_mail=False, stop_after_commit=False):
        events = []
        test = self
        class FakeOutlook:
            def prepare_forward(self, reference, attachment, settings):
                events.append(("prepared", attachment.name))
                test.assertTrue(attachment.name.startswith("01_"))
                return SimpleNamespace(Close=lambda _: None)

            def deliver(self, forwarded, settings):
                events.append(("mail", "attempted"))
                if fail_mail:
                    raise RuntimeError("ambiguous send")
                return "sent" if settings["send_immediately"] else "draft_opened"

            def complete(self, reference, scan):
                events.append(("complete", reference.entry_id))
        backend = sql or FakeSql()
        def database_factory(path):
            database = DailyDatabase(path, backend)
            database.schema = lambda: []
            original = database.import_rows
            def imported(rows, replace, job):
                events.append(("database", rows))
                result = original(rows, replace, job)
                if stop_after_commit:
                    service.stop.set()
                return result
            database.import_rows = imported
            return database
        service = PerformanceService(self.config, threading.Event(),
            outlook=FakeOutlook(), database_factory=database_factory, processor=self.artifacts)
        return service, backend, events

    def test_only_daily_routes_01_mail_and_05_database_and_deduplicates_restart(self):
        service, backend, events = self.service()
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "success")
        self.assertEqual(backend.commits, 1)
        self.assertEqual(next(e[1] for e in events if e[0] == "database"), [record()])
        self.assertTrue(all("PFM_ADJUST" not in s for s in backend.statements))
        self.assertEqual([e[0] for e in events], ["prepared", "database", "mail", "complete"])
        outputs = [Path(p).name for p in result["result"]["outputs"]]
        self.assertEqual(set(outputs), {"01_source_Oct26.xlsx", "05_source_Oct26.xlsx"})
        self.assertEqual(result["result"]["mail_stage"], "01")
        self.assertEqual(result["result"]["database_stage"], "05")
        restarted, backend2, events2 = self.service()
        duplicate = replace(self.reference, entry_id="forwarded_duplicate")
        self.assertIsNone(restarted.process_message(duplicate, self.source, self.legacy, restarted.outlook))
        self.assertEqual(events2, [])
        self.assertEqual(backend2.commits, 0)

    def test_database_failure_never_sends_or_publishes(self):
        service, backend, events = self.service(FakeSql(fail_insert=True))
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(backend.rollbacks, 1)
        self.assertFalse(any(e[0] == "mail" for e in events))
        self.assertFalse(Path(self.config["output"]).exists())
        self.assertEqual(service.ledger.data["files"][digest(self.source)]["database"], "not_saved")

    def test_post_commit_readback_failure_blocks_send_and_replay(self):
        service, backend, events = self.service(FakeSql(bad_readback=True))
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(backend.commits, 1)
        self.assertEqual(backend.rollbacks, 0)
        self.assertFalse(any(e[0] == "mail" for e in events))
        self.assertEqual(service.ledger.data["files"][digest(self.source)]["database"], "committed_unverified")
        self.assertIsNone(service.process_message(self.reference, self.source, self.legacy, service.outlook))
        self.assertEqual(backend.commits, 1)

    def test_ambiguous_send_preserves_sql_and_never_resends(self):
        service, backend, events = self.service(fail_mail=True)
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "failed")
        state = service.ledger.data["files"][digest(self.source)]
        self.assertEqual((state["database"], state["mail"]), ("verified", "unknown"))
        self.assertIsNone(service.process_message(self.reference, self.source, self.legacy, service.outlook))
        self.assertEqual(sum(e[0] == "mail" for e in events), 1)
        self.assertEqual(backend.commits, 1)
        self.assertTrue(self.source.exists())

    def test_cancel_after_commit_preserves_sql_without_sending(self):
        service, backend, events = self.service(stop_after_commit=True)
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "cancelled")
        self.assertEqual(backend.commits, 1)
        self.assertFalse(any(e[0] == "mail" for e in events))

    def test_cutoff_mismatch_blocks_all_external_effects(self):
        service, backend, events = self.service()
        reference = replace(self.reference, match={**self.reference.match, "mail_date": "28/10/2026"})
        result = service.process_message(reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(events, [])
        self.assertEqual(backend.commits, 0)

    def test_wrong_cutoff_mail_does_not_block_the_correct_mail_for_the_same_file(self):
        service,backend,events=self.service()
        wrong=replace(self.reference,entry_id="wrong",match={**self.reference.match,"mail_date":"28/10/2026"})
        self.assertEqual(service.process_message(wrong,self.source,self.legacy,service.outlook)["status"],"failed")
        self.assertEqual(service.ledger.data["files"],{})
        self.assertEqual(service.process_message(self.reference,self.source,self.legacy,service.outlook)["status"],"success")
        self.assertEqual(backend.commits,1)

    def test_draft_is_not_recorded_as_sent_or_archived(self):
        service, backend, events = self.service()
        self.legacy["mail"]["send_immediately"] = False
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "draft_opened")
        self.assertTrue(self.source.exists())
        self.assertFalse(any(e[0] == "complete" for e in events))

    def test_daily_database_rejects_adjust_and_replacement(self):
        database = DailyDatabase("unused", FakeSql()).__enter__()
        with self.assertRaises(ValidationError):
            database.import_rows([replace(record(), report="adjust")], False, self.job)
        with self.assertRaises(ValidationError):
            database.import_rows([record()], True, self.job)

    def test_existing_scope_is_not_overwritten(self):
        fields, values = DailyDatabase.payload(record())
        service, backend, events = self.service(FakeSql(initial=[dict(zip(fields, values))]))
        result = service.process_message(self.reference, self.source, self.legacy, service.outlook)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(backend.commits, 0)
        self.assertFalse(any(s.startswith("DELETE") for s in backend.statements))
        self.assertFalse(any(e[0] == "mail" for e in events))

    def test_encrypted_attachment_contains_only_01(self):
        self.legacy["mail"].update(zip_attachment=True, zip_password="test-only-password")
        stage01 = self.job.folder / "01_Daily_Oct26.xlsx"
        stage01.write_bytes(b"fixture stage01")
        path = make_attachment(stage01, self.legacy["mail"], self.job)
        import pyzipper
        with pyzipper.AESZipFile(path) as archive:
            self.assertEqual(archive.namelist(), [stage01.name])
            with self.assertRaises(RuntimeError):
                archive.read(stage01.name)
            archive.setpassword(b"test-only-password")
            self.assertEqual(archive.read(stage01.name), b"fixture stage01")

    def test_other_stages_cannot_be_attached_even_when_stage01_is_missing(self):
        for stage in ('02','05'):
            for zipped in (False,True):
                with self.subTest(stage=stage,zipped=zipped):
                    path=self.job.folder/(stage+'_Daily_Oct26.xlsx')
                    path.write_bytes(b'not the mail stage')
                    with self.assertRaisesRegex(ValidationError,'01'):
                        make_attachment(path,{'zip_attachment':zipped},self.job)
                    self.assertFalse(path.with_suffix('.zip').exists())

    def test_process_requires_stage01_instead_of_falling_back_to_stage02(self):
        artifacts=self.artifacts(None,None,self.job)
        artifacts.mail_file.unlink()
        (self.job.folder/'02_source_Oct26.xlsx').write_bytes(b'wrong stage')
        with patch('routine.daily.run_daily',return_value=([record()],[str(artifacts.database_file)])):
            with self.assertRaisesRegex(ValidationError,'01/05'):
                process_daily(Path('source_Oct26.xlsx'),self.settings,self.job)

    def test_changed_stage01_blocks_database_and_mail(self):
        service,backend,events=self.service()
        def changed(source,settings,job):
            artifacts=self.artifacts(source,settings,job)
            artifacts.mail_file.write_bytes(b'changed stage01')
            return artifacts
        service.processor=changed
        result=service.process_message(self.reference,self.source,self.legacy,service.outlook)
        self.assertEqual(result['status'],'failed')
        self.assertEqual(backend.commits,0)
        self.assertEqual(events,[])

    def test_unsafe_zip_does_not_write_outside_process(self):
        archive = self.root / "unsafe.zip"
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr("../outside.txt", "bad")
        with self.assertRaises(ValidationError):
            extract_zip(archive, self.root / "extract")
        self.assertFalse((self.root / "outside.txt").exists())

    def test_file_match_blocks_ambiguous_dates(self):
        scan = {"network_folder": str(self.root), "require_exact_date_match": False,
                "network_file_regex": r"^(?P<file_date>\d{6})_A1015_A1015_Daily_Recovery_Performance_(?P<month_token>[A-Za-z]{3}\d{2})\.zip(?:\.(?:gpg|pgp))?$"}
        self.assertEqual(find_archive(scan, self.reference.match), self.source)
        (self.root / self.source.name.replace("261028", "261029")).write_bytes(b"other")
        with self.assertRaises(ValidationError):
            find_archive(scan, self.reference.match)
        self.assertEqual(find_archive(scan, self.reference.match, self.source.name), self.source)

    def test_mode_is_always_daily_even_though_upstream_supports_adjust(self):
        artifacts = self.artifacts(None, None, self.job)
        source = Path("source_Oct26.xlsx")
        with patch("routine.daily.run_daily", return_value=([record()], [str(artifacts.database_file)])) as engine, \
             patch("routine.daily.records_from_05", return_value=(record(),)):
            result = process_daily(source, self.settings, self.job)
        self.assertEqual(engine.call_args.args[1], "daily")
        self.assertNotIn("BAYKAMC", engine.call_args.args[2])
        self.assertEqual(result.mail_file.name, "01_source_Oct26.xlsx")

    def test_daily_keeps_oa15_zero_while_adjust_rejects_it(self):
        grid = [[None] * 12 for _ in range(6)]
        grid[0][0] = "Bucket WO0"
        grid[1][7:11] = ["EFF.", "STANDARD", "SCORE", "RANK"]
        grid[2][1:3] = ["NO.", "NAME"]
        grid[3][1:8] = [1, "OA 15", 200, 1000, "-", None, "-"]
        grid[4][1:8] = [2, "OA 16", 200, 1000, 20, None, 2]
        grid[5][2] = "TOTAL"
        removed, edits = process_plan(grid, "daily")
        self.assertEqual(removed, [])
        self.assertEqual(next(row for row in edits.values() if row[2] == "OA 15")[5], 0)
        self.assertEqual(process_plan(grid, "adjust")[0], list(range(6)))


class Saved05Tests(unittest.TestCase):
    def cells(self):
        values = {"B20": 1, "C20": "OA 15", "D20": 200, "E20": 1000,
                  "F20": 50, "G20": 5, "H20": "FC", "I20": "B01",
                  "J20": "2026-10-27", "K20": "2026-10-29", "L20": "OCT-26",
                  "C21": "TOTAL", "D21": 200}
        return {ref: {"value": value} for ref, value in values.items()}

    def test_import_reads_saved_stage05_metadata_and_ignores_total(self):
        package = SimpleNamespace(sheets={"FC": "", "SummaryByOACode": ""}, cells=lambda sheet: self.cells())
        with patch("routine.daily.Package", return_value=package):
            self.assertEqual(records_from_05("05_Daily_Oct26.xlsx", ("FC",)), (record(),))

    def test_missing_metadata_blocks_import(self):
        cells = self.cells()
        del cells["I20"]
        package = SimpleNamespace(sheets={"FC": "", "SummaryByOACode": ""}, cells=lambda sheet: cells)
        with patch("routine.daily.Package", return_value=package):
            with self.assertRaises(ValidationError):
                records_from_05("05_Daily_Oct26.xlsx", ("FC",))

    def test_mail_stages_cannot_be_used_as_database_input(self):
        for stage in ("01", "02"):
            with self.subTest(stage=stage), self.assertRaises(ValidationError):
                records_from_05(stage + "_Daily_Oct26.xlsx", ("FC",))




class GpgCompatibilityTests(unittest.TestCase):
    def exercise(self, loopback):
        from routine.files import unpack_archive
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "input.zip.gpg"
            archive.write_bytes(b"fixture encrypted bytes")
            job = SimpleNamespace(folder=root / "work", check=lambda: None)
            job.folder.mkdir()
            options = b"--passphrase-fd\n" + (b"--pinentry-mode\n" if loopback else b"")
            commands = []
            def run(command, **kwargs):
                commands.append((command, kwargs))
                if "--dump-options" in command:
                    return SimpleNamespace(returncode=0, stdout=options)
                destination = Path(command[command.index("--output") + 1])
                with zipfile.ZipFile(destination, "w") as zipped:
                    zipped.writestr("Daily_Oct26.xlsm", b"workbook fixture")
                self.assertEqual(kwargs["input"], b"fixture-only-passphrase\n")
                self.assertNotIn("fixture-only-passphrase", command)
                return SimpleNamespace(returncode=0, stdout=b"")
            with patch("routine.files.subprocess.run", side_effect=run):
                output = unpack_archive(archive, job,
                    {"enabled": True, "gpg_path": "fixture-gpg.exe", "passphrase": "fixture-only-passphrase"},
                    r"Daily_Oct26\.xlsm")
            self.assertTrue(output.is_file())
            self.assertEqual("--pinentry-mode" in commands[1][0], loopback)

    def test_gpg20_does_not_receive_unsupported_pinentry_flag(self):
        self.exercise(False)

    def test_modern_gpg_uses_loopback(self):
        self.exercise(True)



class MonitorPreflightTests(unittest.TestCase):
    def test_unavailable_database_blocks_scanning_and_processing(self):
        from contextlib import nullcontext
        from routine.service import monitor
        with tempfile.TemporaryDirectory() as folder:
            config = default_config()
            config.update(logs=str(Path(folder) / "Performance"), output=str(Path(folder) / "Completed"))
            with patch("routine.service.apartment", return_value=nullcontext()), \
                 patch("routine.service.DailyDatabase", side_effect=RuntimeError("provider unavailable")), \
                 patch("routine.service.Outlook") as outlook:
                with self.assertRaisesRegex(ValidationError, "ODBC/OLE DB"):
                    monitor(config, threading.Event(), once=True)
                self.assertFalse(outlook.called)
            self.assertFalse((Path(config["logs"]) / "Logs" / "state.json").exists())

if __name__ == "__main__":
    unittest.main()
