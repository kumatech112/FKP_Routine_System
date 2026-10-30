"""Isolated native UI checks; no Outlook, SQL writes or production monitor."""
from pathlib import Path
from unittest.mock import patch
import argparse
import json
import tkinter as tk

from routine.config import default_config
from routine.ui import Application
from routine.vendor.fkp.persistence import dump


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--destination", type=Path, required=True)
    args = parser.parse_args()
    target = args.destination.resolve()
    target.mkdir(parents=True, exist_ok=False)
    config = default_config()
    config.update(logs=str(target / "Performance"), output=str(target / "Completed"),
                  geometry="1180x780+20+20")
    fake_mail = {"mail": {"send_from": "fixture@example.invalid", "to": ["fixture-group"],
                         "cc": [], "send_immediately": True, "zip_attachment": True}}
    with patch("routine.ui.load_legacy", return_value=fake_mail), patch("routine.ui.monitor") as monitor:
        app = Application(config, target / "settings.local.json")
        app.update()
        assert not monitor.called
        assert [app.notebook.tab(i, "text") for i in range(3)] == ["หน้าหลัก", "ประวัติการทำงาน", "ตั้งค่า"]
        sample = {"id": "fixture", "time": "2026-10-29T10:20:30+07:00",
                  "task": "Performance Daily", "status": "success", "duration": 12.3,
                  "details": "ตรวจสอบขั้น 01 และ 05 แล้ว", "evidence": str(target)}
        dump(Path(config["logs"]) / "Logs" / "History" / "fixture.json", sample)
        app.refresh_history()
        assert len(app.history_table.tree.get_children()) == 1
        app.search.set("ไม่พบ")
        assert not app.filtered
        app.search.set("Daily")
        app.history_table.tree.selection_set("fixture")
        app.sort_history("duration")
        assert app.history_table.tree.selection() == ("fixture",)
        from PIL import ImageGrab
        images = []
        for width, height, name in ((1180, 780, "normal"), (1000, 650, "constrained")):
            app.geometry(f"{width}x{height}+20+20")
            for index, page in enumerate(("main", "history", "settings")):
                app.notebook.select(index)
                app.update()
                filename = target / f"{name}_{page}.png"
                ImageGrab.grab((app.winfo_rootx(), app.winfo_rooty(),
                                app.winfo_rootx()+app.winfo_width(),
                                app.winfo_rooty()+app.winfo_height())).save(filename)
                images.append(str(filename))
        app.fields["logs"].set(str(target / "Changed"))
        app.start()
        assert not monitor.called and app.worker is None
        app.reset()
        assert app.dirty()
        app.destroy()
        reopened = Application(config, target / "settings.local.json")
        reopened.save()
        assert (target / "settings.local.json").is_file()
        assert json.loads((target / "settings.local.json").read_text(encoding="utf-8"))["logs"] == config["logs"]
        reopened.destroy()
        dump(target / "verification.json", {"status": "verified", "screenshots": images,
             "production_monitor_started": False, "database_writes": 0, "emails_sent": 0,
             "native_pages": 3, "search_sort_selection": True, "settings_persistence": True})


if __name__ == "__main__":
    main()
