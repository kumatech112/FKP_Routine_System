"""Multi-workflow routing, user isolation, OCR completeness and Payment partial transactions."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import copy
import json
import re
import tempfile
import threading
import unittest
import zipfile

from routine.config import engine_settings
from routine.delivery import package_files
from routine.enforcement import FIELDS, approve_review, table_grids
from routine.matching import body_names, locate_body_file
from routine.outlook import MailRef
from routine.payment import ActionBatch, PaymentArtifacts, commit_actions
from routine.service import RoutineJob
from routine.settings import TASKS, default_settings, load, runtime_settings, save, task_config
from routine.vendor.fkp.domain import ValidationError
from routine.vendor.fkp.persistence import digest
from routine.workflows import FileWorkflowService


class MultiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.values = default_settings()
        self.values.update(logs_root=str(self.root/"Logs"), output_root=str(self.root/"Completed"))
        self.values["secrets"] = {"passphrase": "fixture-pass", "zip_password": "fixture-zip"}
        self.values["archive"]["archive_processed_files"] = False

    def tearDown(self):
        self.temp.cleanup()

    def test_four_tasks_and_new_payment_path(self):
        self.assertEqual(set(self.values["workflows"]), set(TASKS))
        self.assertTrue(self.values["payment_database"].endswith("FKP Payment Online Round.mdb"))
        for task in TASKS:
            cfg = task_config(self.values, task)
            self.assertTrue(cfg["logs"].endswith(TASKS[task]))
            self.assertNotIn("KUMA BOT", cfg["logs"])

    def test_profile_roundtrip_is_encrypted_and_does_not_write_install_folder(self):
        path = self.root/"user-profile"/"settings.json"
        save(self.values, path)
        raw = path.read_text(encoding="utf-8")
        self.assertNotIn("fixture-pass", raw)
        self.assertNotIn("fixture-zip", raw)
        self.assertEqual(load(path)["secrets"], self.values["secrets"])
        self.assertNotIn("secrets", json.loads(raw))

    def test_another_windows_identity_does_not_get_silent_blank_secrets(self):
        path = self.root/"settings.json"
        save(self.values, path)
        with patch("routine.settings.unseal", side_effect=RuntimeError("other user")):
            with self.assertRaisesRegex(ValidationError, "ผู้ใช้อื่น"):
                load(path)

    def test_documents_ignore_subject_and_sender_but_require_a1015_body(self):
        runtime = runtime_settings(self.values, "customer_documents")
        self.assertEqual(runtime["scan"]["subject_regex"], ".*")
        self.assertEqual(runtime["scan"]["sender_contains"], [])
        body = "wrong subject: A1015L_Document_Round2_22102026.zip.gpg and A1015_Document_Round3_22102026.zip.pgp"
        self.assertEqual(len(body_names(body, "customer_documents")), 2)
        self.assertEqual(body_names("A1099L_Document_Round2_22102026.zip.gpg", "customer_documents"), [])

    def test_body_name_can_omit_transport_date_but_never_pick_ambiguous_file(self):
        leaf = "A1015L_Document_Round2_22102026.zip.gpg"
        source = self.root/("261022_"+leaf)
        source.write_bytes(b"one")
        self.assertEqual(locate_body_file(self.root, leaf), source)
        (self.root/("261023_"+leaf)).write_bytes(b"two")
        with self.assertRaises(ValidationError):
            locate_body_file(self.root, leaf)

    def test_enforcement_mail_requires_its_own_recipients(self):
        runtime = runtime_settings(self.values, "enforcement")
        self.assertTrue(runtime["mail"]["enabled"])
        self.assertEqual(runtime["mail"]["to"], [])
        name = "261028_A1015L_A1015L_Approval_Enforcement_28102026.zip.gpg"
        self.assertEqual(body_names(name, "enforcement"), [name])

    def test_document_zip_preserves_nested_file_names_and_empty_password_is_blocked(self):
        cfg = engine_settings(task_config(self.values, "customer_documents"))
        job = RoutineJob(self.root/"job", cfg)
        tree = job.folder/"Extracted"
        files = [tree/"one"/"same.pdf", tree/"two"/"same.pdf"]
        for i, file in enumerate(files):
            file.parent.mkdir(parents=True)
            file.write_bytes(str(i).encode())
        mail = runtime_settings(self.values, "customer_documents")["mail"]
        result = package_files(files, mail, job, "documents", relative_root=tree)
        import pyzipper
        with pyzipper.AESZipFile(result) as archive:
            archive.setpassword(b"fixture-zip")
            self.assertEqual(archive.namelist(), ["one/same.pdf", "two/same.pdf"])
        with self.assertRaises(ValidationError):
            package_files(files, {**mail, "zip_password": ""}, job, "invalid", relative_root=tree)

    def test_table_detection_includes_the_blank_last_column(self):
        from PIL import Image, ImageDraw
        image = Image.new("RGB", (1100, 500), "white")
        draw = ImageDraw.Draw(image)
        xs = [30+i*45 for i in range(22)]
        ys = [150, 260, 310, 360, 410]
        for y in ys:
            draw.line((xs[0],y,xs[-1],y), fill="black", width=2)
        for x in xs[:-1]:
            draw.line((x,ys[0],x,ys[-1]), fill="black", width=2)
        tables = table_grids(image)
        self.assertEqual([(len(t["x"])-1,len(t["y"])-2) for t in tables], [(21,3)])

    def test_ocr_never_publishes_before_every_cell_including_blank_is_reviewed(self):
        cfg = engine_settings(task_config(self.values, "enforcement"))
        job = RoutineJob(self.root/"job", cfg)
        data = {"status":"needs_review","tables":[{"headers":list(FIELDS)}],
                "cells":[{"id":"blank","state":"needs_review","value":""}]}
        path = job.folder/"review.json"
        path.write_text(json.dumps(data),encoding="utf-8")
        with patch("routine.enforcement.publish_outputs") as publish:
            with self.assertRaisesRegex(ValidationError,"ยังมีช่อง"):
                approve_review(path, {}, job, cfg)
            self.assertFalse(publish.called)

    def test_body_multiple_files_are_independent_units_in_one_mail(self):
        flow = self.values["workflows"]["customer_documents"]
        flow["network_folder"] = str(self.root/"inbox")
        Path(flow["network_folder"]).mkdir()
        names = ["A1015L_Document_Round2_22102026.zip.gpg","A1015_Document_Round3_22102026.zip.pgp"]
        for i,name in enumerate(names):
            (Path(flow["network_folder"])/name).write_bytes(f"encrypted-fixture-{i}".encode())
        reference = MailRef("mail","store","wrong subject"," ".join(names),{})
        deliveries=[]
        class FakeOutlook:
            def scan_workflow(self, settings): return [reference]
            def prepare_forward(self, reference, attachment, settings):
                return SimpleNamespace(Close=lambda _:None)
            def deliver(self, forwarded, settings):
                deliveries.append("sent");return "sent"
            def complete(self,*_): pass
        def extracted(source, job, decrypt):
            target=job.folder/"Extracted"
            target.mkdir()
            (target/"document.pdf").write_bytes(b"fixture document")
            return target
        service=FileWorkflowService(self.values,"customer_documents",threading.Event(),outlook=FakeOutlook())
        with patch("routine.workflows.extract_payload",side_effect=extracted):
            result=service.poll()
        self.assertEqual(len(result),2)
        self.assertEqual(deliveries,["sent","sent"])
        self.assertEqual(len(service.ledger.data["files"]),2)
        self.assertEqual(service.poll(),[])


class ActionSql:
    def __init__(self, fail=False):
        self.rows=[];self.before=[];self.fail=fail;self.commits=0;self.rollbacks=0
    def run(self,sql,args=(),rows=True):
        if "TOP (0)" in sql:return []
        if sql.startswith("SELECT"):return copy.deepcopy(self.rows)
        if sql.startswith("DELETE"):self.rows=[];return []
        if sql.startswith("INSERT"):
            if self.fail:raise RuntimeError("second database failure")
            fields=re.findall(r"\[([^]]+)\]",sql.split("VALUES")[0])[2:]
            self.rows.append(dict(zip(fields,args)));return []
        raise AssertionError(sql)
    def begin(self):self.before=copy.deepcopy(self.rows)
    def commit(self):self.commits+=1
    def rollback(self):self.rows=self.before;self.rollbacks+=1
    def close(self):pass


class PaymentTransactions(unittest.TestCase):
    def test_partial_database_success_is_preserved_and_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            values=default_settings();values.update(logs_root=str(root/"Logs"),output_root=str(root/"Completed"))
            job=RoutineJob(root/"job",engine_settings(task_config(values,"payment_online")))
            output=job.folder/"payment.xlsx";output.write_bytes(b"fixture")
            row={"AgreementN":"A1","COLLCETION":"Payment Round 2 fixture","type":"DATA"}
            batches=(ActionBatch("FC","[dbo].[Calls]",tuple(row),(row,),"one"),
                     ActionBatch("KHB","[dbo].[Calls]",tuple(row),(row,),"two"))
            artifacts=PaymentArtifacts(output,batches,1,digest(output))
            first,second=ActionSql(),ActionSql(fail=True)
            with self.assertRaises(RuntimeError):
                commit_actions(artifacts,job,lambda key:{"one":first,"two":second}[key])
            self.assertEqual((first.commits,first.rollbacks),(1,0))
            self.assertEqual((second.commits,second.rollbacks),(0,1))
            self.assertEqual(job.database_outcome,"partial_verified")
            states=json.loads((job.evidence/"payment_groups.json").read_text())
            self.assertEqual([s["status"] for s in states],["verified","not_saved"])


if __name__=="__main__":
    unittest.main()
