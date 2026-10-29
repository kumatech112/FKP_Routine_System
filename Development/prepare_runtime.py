"""Stage redistributable executables/libraries only; never copy keys or private settings."""
from pathlib import Path
import shutil,json

root=Path(__file__).resolve().parent
target=root/"runtime"
target.mkdir(exist_ok=True)
records=[]
gpg_tools={"gpg2.exe","gpg-agent.exe","gpg-connect-agent.exe","gpgconf.exe","gpgv2.exe",
           "gpgme-w32spawn.exe","gpg-protect-tool.exe"}
for name,source in [("gpg",Path(r"D:\Documents\Routine Bot System\gpg")),
                    ("ocr",Path(r"C:\Program Files\Tesseract-OCR"))]:
    output=target/name
    output.mkdir(exist_ok=True)
    for path in source.iterdir():
        allowed=(path.suffix.lower()==".dll"
                 or name=="gpg" and (path.name in gpg_tools or path.name.startswith(("pinentry","gpg2keys_")))
                 or name=="ocr" and path.name=="tesseract.exe"
                 or path.name.lower().startswith(("copying","license")))
        if path.is_file() and allowed:
            shutil.copy2(path,output/path.name)
            records.append(str((output/path.name).relative_to(root)))
    if name=="ocr":
        data=output/"tessdata"
        data.mkdir(exist_ok=True)
        for lang in ("tha","eng","osd"):
            shutil.copy2(source/"tessdata"/(lang+".traineddata"),data/(lang+".traineddata"))
            records.append(str((data/(lang+".traineddata")).relative_to(root)))
# UnRAR is the separately redistributable extraction component (License.txt section 3a).
unrar=target/"unrar"
unrar.mkdir(exist_ok=True)
for filename in ("UnRAR.exe","License.txt"):
    shutil.copy2(Path(r"C:\Program Files\WinRAR")/filename,unrar/filename)
    records.append(str((unrar/filename).relative_to(root)))
(root/"runtime_manifest.json").write_text(json.dumps({"files":records},indent=2),encoding="utf-8")
print({"runtime_files":len(records)})
