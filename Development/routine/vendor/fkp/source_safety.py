"""Disable open-time external refresh in disposable OOXML copies only."""
from pathlib import Path
from zipfile import ZipFile
from xml.dom import minidom
import os
import uuid

from .persistence import replace_file


def disable_refresh(path):
    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return
    changed = {}
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if name == "xl/connections.xml" or name.startswith(("xl/pivotCache/pivotCacheDefinition", "xl/queryTables/")) and name.endswith(".xml"):
                document = minidom.parseString(archive.read(name))
                for element in document.getElementsByTagName("*"):
                    for attribute in ("refreshOnLoad", "refreshOnFileOpen", "backgroundRefresh", "backgroundQuery"):
                        if element.hasAttribute(attribute):
                            element.setAttribute(attribute, "0")
                changed[name] = document.toxml(encoding="utf-8")
                document.unlink()
        if not changed:
            return
        temporary = path.with_name(uuid.uuid4().hex + path.suffix)
        with temporary.open("xb") as stream, ZipFile(stream, "w") as output:
            for info in archive.infolist():
                output.writestr(info, changed.get(info.filename, archive.read(info.filename)))
    replace_file(temporary, path)
