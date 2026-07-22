#!/usr/bin/env python3

from pathlib import Path
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

if len(sys.argv) < 2:
    print('Usage: python3 remove_remaining_layout_junk.py "file.docx"')
    sys.exit(1)

infile = Path(sys.argv[1])
outfile = infile.with_name(infile.stem + "_NO_REMAINING_LAYOUT_JUNK.docx")

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", NS_W)

REMOVE = {
    "br",
    "footnoteReference",
    "endnoteReference",
    "sectPr",
    "lastRenderedPageBreak",
}

def localname(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag

def clean(elem, counter):
    i = 0
    while i < len(elem):
        child = elem[i]
        name = localname(child.tag)
        if name in REMOVE:
            counter[name] += 1
            elem.remove(child)
            continue
        clean(child, counter)
        i += 1

def should_process(name):
    return (
        name == "word/document.xml"
        or name == "word/footnotes.xml"
        or name == "word/endnotes.xml"
        or name.startswith("word/header")
        or name.startswith("word/footer")
    ) and name.endswith(".xml")

with zipfile.ZipFile(infile, "r") as zin:
    files = {name: zin.read(name) for name in zin.namelist()}

counter = Counter()

for name, data in list(files.items()):
    if not should_process(name):
        continue
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        continue
    clean(root, counter)
    files[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

with zipfile.ZipFile(outfile, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        zout.writestr(name, data)

print("Wrote:", outfile)
print("Removed:")
for k, v in sorted(counter.items()):
    print(f"  {k}: {v}")
