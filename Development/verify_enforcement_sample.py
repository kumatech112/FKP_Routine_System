from pathlib import Path
from routine.runtime import bootstrap_windows
bootstrap_windows()
from routine.config import default_config,engine_settings
from routine.service import RoutineJob
from routine.enforcement import prepare_enforcement
root=Path(r'\\10.0.0.97\kuma$\Z\Projects\FKP_Routine_System\Archive\enforcement_verification')
root.mkdir(exist_ok=False)
c=default_config();c['logs']=str(root/'Logs');c['output']=str(root/'Completed')
j=RoutineJob(root/'Process',engine_settings(c))
r=prepare_enforcement(Path(r'D:\Downloads\261028_A1015L_A1015L_Approval_Enforcement_28102026'),j,r'C:\Program Files\Tesseract-OCR\tesseract.exe')
print({'pages':r['page_count'],'tables':len(r['tables']),'cells':r['cell_count'],'rows':[x['rows'] for x in r['tables']],'status':r['status']})
