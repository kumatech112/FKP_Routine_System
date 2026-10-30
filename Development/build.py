"""Build a portable, asInvoker Windows release without enabling monitoring."""
from pathlib import Path
import importlib.metadata
import json
import os
import shutil
import subprocess
import sys
from routine.runtime import bootstrap_windows
from routine import __version__

ROOT=Path(__file__).resolve().parents[1]
DEV=ROOT/"Development"
RELEASE=ROOT/"Program"/("FKP_Routine_System_"+__version__)
WORK=ROOT/"Archive"/("build_"+__version__)


def main():
    bootstrap_windows()
    if RELEASE.exists():
        raise SystemExit("Release already exists; preserve it before rebuilding.")
    WORK.mkdir(parents=True,exist_ok=True)
    command=[sys.executable,"-m","PyInstaller","--noconfirm","--clean","--onedir","--windowed",
        "--name",RELEASE.name,"--distpath",str(RELEASE.parent),"--workpath",str(WORK/"work"),
        "--specpath",str(WORK),"--paths",str(DEV),
        "--add-data",str(DEV/"Templates")+";Templates",
        "--add-data",str(DEV/"assets")+";assets",
        "--add-data",str(DEV/"runtime")+";runtime",
        "--add-data",str(DEV/"legacy_defaults.json")+";.",
        "--hidden-import","win32timezone","--hidden-import","pythoncom",
        "--hidden-import","win32com.client","--hidden-import","win32crypt",
        "--hidden-import","pyzipper","--icon",str(DEV/"assets"/"app.ico")]
    for module in ("torch","tensorflow","paddle","paddleocr","scipy","matplotlib","pandas"):
        command += ["--exclude-module",module]
    command += [str(DEV/"main.py")]
    environment=dict(os.environ)
    environment["PYINSTALLER_CONFIG_DIR"]=str(WORK/"cache")
    subprocess.run(command,cwd=DEV,check=True,env=environment)
    (RELEASE/(RELEASE.name+".exe")).rename(RELEASE/"FKP_Routine_System.exe")
    licenses=RELEASE/"licenses"
    licenses.mkdir()
    source=Path(sys.base_prefix)/"LICENSE.txt"
    if source.is_file():shutil.copyfile(source,licenses/"Python.txt")
    packages={}
    for package in ("pywin32","pyzipper","pycryptodomex","pyinstaller","openpyxl","Pillow","numpy",
                    "opencv-python","pymupdf","pytesseract","rarfile","pycryptodome","cryptography"):
        distribution=importlib.metadata.distribution(package)
        packages[package]=distribution.version
        for index,filename in enumerate(distribution.files or []):
            if any(part.lower().startswith(("license","copying")) for part in filename.parts):
                original=Path(distribution.locate_file(filename))
                if original.is_file():shutil.copyfile(original,licenses/f"{package}_{index}_{original.name}")
    extra=DEV/"third_party_licenses"
    if extra.exists():
        for path in extra.rglob("*"):
            if path.is_file():
                target=licenses/path.relative_to(extra)
                target.parent.mkdir(parents=True,exist_ok=True)
                shutil.copyfile(path,target)
    revision=subprocess.run(["git","-c","safe.directory="+ROOT.as_posix(),"-C",str(ROOT),
                             "rev-parse","--verify","HEAD"],text=True,capture_output=True)
    source_commit=revision.stdout.strip() if revision.returncode==0 else None
    (RELEASE/"build.json").write_text(json.dumps({
        "version":__version__,"source_commit":source_commit,
        "commit_status":"recorded" if source_commit else "incomplete: no Git commit is available",
        "workflows":["payment_online","performance_daily","customer_documents","enforcement","performance_success_rate"],
        "mail_stage":"01","daily_database_stage":"05","profile_storage":"Windows user DPAPI",
        "requires_administrator":False,"dependencies":packages,
    },indent=2),encoding="utf-8")


if __name__=="__main__":main()
