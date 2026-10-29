"""FKP Routine System entry point. Verification never sends mail or writes production SQL."""
import argparse
from pathlib import Path
import threading

from routine.runtime import bootstrap_windows
from routine.settings import load, settings_path, task_config
from routine.config import engine_settings
from routine.daily import process_daily
from routine.payment import prepare_payment
from routine.enforcement import prepare_enforcement
from routine.service import RoutineJob
from routine.workflows import run
from routine.vendor.fkp.office import apartment
from routine.vendor.fkp.persistence import dump


def main():
    bootstrap_windows()
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", type=Path)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--once", action="store_true")
    mode.add_argument("--monitor", action="store_true")
    mode.add_argument("--verify-source", type=Path)
    mode.add_argument("--verify-payment", type=Path)
    mode.add_argument("--verify-enforcement", type=Path)
    parser.add_argument("--verification-root", type=Path)
    args = parser.parse_args()
    values = load(args.config)
    if args.verify_source or args.verify_payment or args.verify_enforcement:
        if not args.verification_root:
            parser.error("verification requires --verification-root")
        target = args.verification_root.resolve()
        target.mkdir(parents=True, exist_ok=False)
        task = "performance" if args.verify_source else ("payment_online" if args.verify_payment else "enforcement")
        config = task_config(values, task)
        config.update(logs=str(target / "Logs"), output=str(target / "Completed"))
        settings = engine_settings(config)
        job = RoutineJob(target / "Process", settings)
        with apartment():
            if args.verify_source:
                result = process_daily(args.verify_source, settings, job)
                result.verify()
                outcome = {"records": len(result.records), "mail_stage": str(result.mail_file),
                           "database_stage": str(result.database_file)}
            elif args.verify_payment:
                result = prepare_payment(args.verify_payment, values["payment_database"], job, run_actions=True)
                outcome = {"records": result.source_rows, "action_rows": sum(len(b.rows) for b in result.batches),
                           "output": str(result.output)}
            else:
                result = prepare_enforcement(args.verify_enforcement, job, values["ocr_path"])
                outcome = {"pages": result["page_count"], "cells": result["cell_count"], "status": result["status"]}
        dump(target / "verification.json", {"task": task, "verified": True,
            "database_writes": 0, "emails_sent": 0, **outcome})
        return 0
    if args.once or args.monitor:
        stop = threading.Event()
        try:
            run(values, stop, once=args.once)
        except KeyboardInterrupt:
            stop.set()
        return 0
    from routine.ui_all import Application
    Application(values, args.config or settings_path()).mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
