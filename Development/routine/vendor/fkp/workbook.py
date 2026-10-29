"""Targeted OOXML edits preserving untouched workbook parts and namespaces."""
from pathlib import Path
from zipfile import ZipFile
from xml.dom import minidom
import copy
import posixpath
import re
import uuid
import xml.etree.ElementTree as ET

from .domain import ValidationError

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": MAIN}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def column(index):
    result = ""
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


def col_index(reference):
    result = 0
    for char in re.match(r"[A-Z]+", reference)[0]:
        result = result * 26 + ord(char) - 64
    return result


def translate(formula, origin, target):
    dx = col_index(target) - col_index(origin)
    dy = int(re.search(r"\d+", target)[0]) - int(re.search(r"\d+", origin)[0])
    def replace(match):
        absolute_col, letters, absolute_row, rn = match.groups()
        return absolute_col + column(col_index(letters) + (0 if absolute_col else dx)) + absolute_row + str(int(rn) + (0 if absolute_row else dy))
    pieces = re.split(r'("(?:[^"]|"")*")', formula)
    return "".join(piece if index % 2 else re.sub(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)", replace, piece) for index, piece in enumerate(pieces))


class Package:
    def __init__(self, path):
        self.path = Path(path)
        with ZipFile(path) as archive:
            self.parts = {part.filename: (part, archive.read(part.filename)) for part in archive.infolist()}
        workbook = ET.fromstring(self.parts["xl/workbook.xml"][1])
        relations = ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"][1])
        targets = {r.attrib["Id"]: posixpath.normpath(posixpath.join("xl", r.attrib["Target"])) if not r.attrib["Target"].startswith("/") else r.attrib["Target"].lstrip("/") for r in relations}
        self.sheets = {s.attrib["name"]: targets[s.attrib[f"{{{REL}}}id"]] for s in workbook.find("s:sheets", NS)}
        self.strings = []
        if "xl/sharedStrings.xml" in self.parts:
            self.strings = ["".join(n.itertext()) for n in ET.fromstring(self.parts["xl/sharedStrings.xml"][1])]

    def cells(self, sheet):
        root = ET.fromstring(self.parts[self.sheets[sheet]][1])
        result = {}
        shared = {}
        for cell in root.findall("s:sheetData/s:row/s:c", NS):
            f = cell.find("s:f", NS)
            if f is not None and f.attrib.get("t") == "shared" and f.text:
                shared[f.attrib["si"]] = (cell.attrib["r"], f.text)
        for cell in root.findall("s:sheetData/s:row/s:c", NS):
            v, f = cell.find("s:v", NS), cell.find("s:f", NS)
            text = v.text if v is not None else None
            kind = cell.attrib.get("t")
            if kind == "s" and text is not None:
                text = self.strings[int(text)]
            elif kind == "inlineStr":
                text = "".join(cell.find("s:is", NS).itertext())
            elif text is not None and kind not in ("str", "e", "b"):
                text = float(text)
            formula = f.text if f is not None else None
            if f is not None and f.attrib.get("t") == "shared" and not formula:
                origin, expression = shared[f.attrib["si"]]
                formula = translate(expression, origin, cell.attrib["r"])
            result[cell.attrib["r"]] = {"value": text, "formula": formula, "type": kind}
        return result

    def clone_sheet(self, source, target):
        if target in self.sheets:
            raise ValidationError("มีชื่อชีตปลายทางนี้แล้ว")
        content_name = "[Content_Types].xml"
        content_doc = minidom.parseString(self.parts[content_name][1])
        content_types = {n.getAttribute("PartName").lstrip("/"): n.getAttribute("ContentType") for n in content_doc.documentElement.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == "Override"}
        mapped = {}
        def clone_part(name):
            if name in mapped:
                return mapped[name]
            directory, base = posixpath.split(name)
            stem, suffix = posixpath.splitext(base)
            new_name = posixpath.join(directory, stem + "_fkp_" + uuid.uuid4().hex[:10] + suffix)
            mapped[name] = new_name
            old_info, old_content = self.parts[name]
            new_info = copy.copy(old_info)
            new_info.filename = new_name
            new_info.orig_filename = new_name
            self.parts[new_name] = (new_info, old_content)
            if name in content_types:
                node = content_doc.createElementNS(content_doc.documentElement.namespaceURI, "Override")
                node.setAttribute("PartName", "/" + new_name)
                node.setAttribute("ContentType", content_types[name])
                content_doc.documentElement.appendChild(node)
            relations_name = posixpath.join(directory, "_rels", base + ".rels")
            if relations_name in self.parts:
                info, content = self.parts[relations_name]
                document = minidom.parseString(content)
                for relationship in document.documentElement.childNodes:
                    if relationship.nodeType != relationship.ELEMENT_NODE or relationship.getAttribute("TargetMode") == "External":
                        continue
                    reference = relationship.getAttribute("Target")
                    resolved = reference.lstrip("/") if reference.startswith("/") else posixpath.normpath(posixpath.join(directory, reference))
                    cloned = clone_part(resolved)
                    relationship.setAttribute("Target", posixpath.relpath(cloned, directory))
                new_relations = posixpath.join(directory, "_rels", posixpath.basename(new_name) + ".rels")
                new_info = copy.copy(info)
                new_info.filename = new_info.orig_filename = new_relations
                self.parts[new_relations] = (new_info, document.toxml(encoding="utf-8"))
                document.unlink()
            return new_name
        cloned = clone_part(self.sheets[source])
        workbook_name, relations_name = "xl/workbook.xml", "xl/_rels/workbook.xml.rels"
        workbook = minidom.parseString(self.parts[workbook_name][1])
        relations = minidom.parseString(self.parts[relations_name][1])
        sheets = workbook.getElementsByTagNameNS(MAIN, "sheets")[0]
        new_id = max(int(n.getAttribute("sheetId")) for n in sheets.childNodes if n.nodeType == n.ELEMENT_NODE) + 1
        relationship_id = "rId_fkp_" + uuid.uuid4().hex[:10]
        node = workbook.createElementNS(MAIN, "sheet")
        node.setAttribute("name", target)
        node.setAttribute("sheetId", str(new_id))
        node.setAttributeNS(REL, "r:id", relationship_id)
        sheets.appendChild(node)
        relationship = relations.createElementNS(relations.documentElement.namespaceURI, "Relationship")
        relationship.setAttribute("Id", relationship_id)
        relationship.setAttribute("Type", REL + "/worksheet")
        relationship.setAttribute("Target", posixpath.relpath(cloned, "xl"))
        relations.documentElement.appendChild(relationship)
        for name, document in ((content_name, content_doc), (workbook_name, workbook), (relations_name, relations)):
            self.parts[name] = (self.parts[name][0], document.toxml(encoding="utf-8"))
            document.unlink()
        self.sheets[target] = cloned

    def patch(self, sheet, changes, caches=None):
        name = self.sheets[sheet]
        info, source = self.parts[name]
        doc = minidom.parseString(source)
        data = doc.getElementsByTagNameNS(MAIN, "sheetData")[0]
        rows = {int(n.getAttribute("r")): n for n in data.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == "row"}
        for address, value in changes.items():
            rn = int(re.search(r"\d+", address)[0])
            if rn not in rows:
                row = doc.createElementNS(MAIN, "row")
                row.setAttribute("r", str(rn))
                following = next((rows[i] for i in sorted(rows) if i > rn), None)
                data.insertBefore(row, following)
                rows[rn] = row
            row = rows[rn]
            cells = {n.getAttribute("r"): n for n in row.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == "c"}
            cell = cells.get(address)
            if cell is None:
                cell = doc.createElementNS(MAIN, "c")
                cell.setAttribute("r", address)
                following = next((n for ref, n in cells.items() if col_index(ref) > col_index(address)), None)
                row.insertBefore(cell, following)
            for node in list(cell.childNodes):
                if node.nodeType == node.ELEMENT_NODE and node.localName in ("f", "v", "is"):
                    cell.removeChild(node)
            if cell.hasAttribute("t"):
                cell.removeAttribute("t")
            if value is None:
                continue
            if isinstance(value, str) and value.startswith("="):
                node = doc.createElementNS(MAIN, "f")
                node.appendChild(doc.createTextNode(value[1:]))
                cell.appendChild(node)
                if caches and address in caches:
                    cached = caches[address]
                    if cached == "":
                        cell.setAttribute("t", "str")
                    node = doc.createElementNS(MAIN, "v")
                    node.appendChild(doc.createTextNode(str(cached)))
                    cell.appendChild(node)
            elif isinstance(value, str):
                cell.setAttribute("t", "inlineStr")
                node = doc.createElementNS(MAIN, "is")
                text = doc.createElementNS(MAIN, "t")
                text.appendChild(doc.createTextNode(value))
                node.appendChild(text)
                cell.appendChild(node)
            else:
                node = doc.createElementNS(MAIN, "v")
                node.appendChild(doc.createTextNode(str(value)))
                cell.appendChild(node)
        self.parts[name] = (info, doc.toxml(encoding="utf-8"))
        doc.unlink()

    def save(self, destination):
        with Path(destination).open("xb") as stream:
            with ZipFile(stream, "w") as archive:
                for info, content in self.parts.values():
                    archive.writestr(info, content)

    def widths(self, sheet):
        root = ET.fromstring(self.parts[self.sheets[sheet]][1])
        result = {}
        for span in root.findall("s:cols/s:col", NS):
            for index in range(int(span.attrib["min"]), int(span.attrib["max"]) + 1):
                result[index] = dict(span.attrib)
        return result

    def widen(self, sheet, widths):
        name = self.sheets[sheet]
        info, source = self.parts[name]
        document = minidom.parseString(source)
        containers = document.getElementsByTagNameNS(MAIN, "cols")
        if containers:
            container = containers[0]
        else:
            container = document.createElementNS(MAIN, "cols")
            document.documentElement.insertBefore(container, document.getElementsByTagNameNS(MAIN, "sheetData")[0])
        existing = self.widths(sheet)
        for child in list(container.childNodes):
            container.removeChild(child)
        spans = []
        for index in sorted(set(existing) | set(widths)):
            attrs = dict(existing.get(index, {}))
            attrs.pop("min", None)
            attrs.pop("max", None)
            if attrs.get("hidden") != "1" and index in widths:
                attrs["width"] = str(max(float(attrs.get("width", 8.43)), widths[index]))
                attrs["customWidth"] = "1"
            if spans and spans[-1][1] + 1 == index and spans[-1][2] == attrs:
                spans[-1][1] = index
            else:
                spans.append([index, index, attrs])
        for first, last, attrs in spans:
            attrs.update(min=str(first), max=str(last))
            node = document.createElementNS(MAIN, "col")
            for attribute, value in attrs.items():
                node.setAttribute(attribute, value)
            container.appendChild(node)
        self.parts[name] = (info, document.toxml(encoding="utf-8"))
        document.unlink()


def source_grid(path, sheet):
    return Package(path).cells(sheet)
