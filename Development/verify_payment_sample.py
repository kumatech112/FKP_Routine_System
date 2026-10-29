from pathlib import Path
from routine.runtime import bootstrap_windows
bootstrap_windows()
from routine.config import default_config,engine_settings
from routine.service import RoutineJob
from routine.vendor.fkp.office import apartment
from routine.payment import prepare_payment
from routine.vendor.fkp.persistence import dump
target=Path(r"\\10.0.0.97\kuma$\Z\Projects\FKP_Routine_System\Archive\payment_verification_02")
target.mkdir(exist_ok=False)
config=default_config();config["logs"]=str(target/"Logs");config["output"]=str(target/"Completed")
job=RoutineJob(target/"Process",engine_settings(config))
with apartment():
    result=prepare_payment(Path(r"\\10.0.0.250\kuma$\Z\KumaBot\20261028_103534_payment_online_round3\A1015_PaymentPostWO_Online_R3_28102026.xlsx"),
        Path(r"\\10.0.0.97\kuma$\X\FKP Payment Online Round.mdb"),job,run_actions=True)
dump(target/"verification.json",{"source_rows":result.source_rows,"actions":[{"name":x.query,"rows":len(x.rows)} for x in result.batches],
    "output":str(result.output),"database_writes":0,"emails_sent":0,"working_mdb_encrypted":not (job.folder/"Payment_work.mdb").exists()})
