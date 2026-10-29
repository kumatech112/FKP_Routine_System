"""Read-only proof: existing local settings decrypt an isolated copy without re-entry."""
from pathlib import Path
import argparse,json
from routine.runtime import bootstrap_windows
from routine.settings import load
from routine.files import extract_payload
from routine.vendor.fkp.persistence import digest
from routine.service import RoutineJob
from routine.config import engine_settings


def main():
    bootstrap_windows()
    parser=argparse.ArgumentParser()
    parser.add_argument('--destination',type=Path,required=True)
    args=parser.parse_args()
    target=args.destination.resolve()
    target.mkdir(parents=True,exist_ok=False)
    project=Path(__file__).resolve().parents[1]
    source=project/'Archive/portable_gpg_verification/Process/261028_A1015_A1015_Daily_Recovery_Performance_Oct26.zip.gpg'
    before=digest(source)
    # An absent isolated profile exercises the existing one-time local migration.
    values=load(target/'Profile/settings.json')
    job=RoutineJob(target/'Process',engine_settings({'logs':str(target/'Logs'),
        'output':str(target/'Completed'),'database':values['database']}))
    extracted=extract_payload(source,job,{'enabled':True,'gpg_path':values['gpg_path'],
        'gpg_home':values['gpg_home'],'passphrase':values['secrets']['passphrase']})
    expected=project/'Archive/portable_gpg_verification/Process/Extracted/Daily_Recovery_Performance_Oct26.xlsm'
    actual=extracted/expected.name
    assert digest(actual)==digest(expected)
    assert digest(source)==before
    report={'verified':True,'reentered_password':False,'existing_local_settings_loaded':True,
        'source_preserved':True,'extracted_files':len(list(extracted.rglob('*.xlsm'))),
        'plaintext_secrets_saved':False,'emails_sent':0,'database_writes':0}
    (target/'verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(json.dumps(report))


if __name__=='__main__':main()
