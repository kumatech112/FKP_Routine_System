"""Read-only real-runtime checks; never scan or send Outlook mail or import SQL."""
from pathlib import Path
import argparse
import json

from routine.runtime import bootstrap_windows
from routine.config import default_config, engine_settings, load_legacy
from routine.daily import DailyDatabase
from routine.files import unpack_archive
from routine.service import RoutineJob
from routine.vendor.fkp.office import apartment
from routine.vendor.fkp.persistence import dump


def main():
    bootstrap_windows()
    parser = argparse.ArgumentParser()
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--archive", type=Path, required=True)
    args = parser.parse_args()
    target = args.destination.resolve()
    target.mkdir(parents=True, exist_ok=False)
    config = default_config()
    config.update(logs=str(target / "Performance"), output=str(target / "Completed"))
    legacy = load_legacy(config)
    settings = engine_settings(config)
    job = RoutineJob(Path(settings["work"]) / "decryption", settings)
    source = unpack_archive(args.archive, job, legacy["decrypt"], legacy["scan"]["excel_regex"])
    result = {"decryption": "verified", "workbook_name": source.name,
              "database_writes": 0, "emails_sent": 0, "source_archived": False}
    try:
        with apartment(), DailyDatabase(settings["database"]) as database:
            schema = database.schema()
            result["schema"] = [{"report": item["report"], "table": item["table"],
                                 "field_count": len(item["fields"])} for item in schema]
            result["database_schema"] = "verified_readonly"
    except Exception as error:
        # Connection exceptions can include credentials, so only retain the class.
        result["database_schema"] = "unverified"
        result["error_type"] = type(error).__name__
    dump(target / "verification.json", result)


if __name__ == "__main__":
    main()
