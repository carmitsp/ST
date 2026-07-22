#!/usr/bin/env python3

from pathlib import Path
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

if len(sys.argv) < 2:
    print("Usage:")
    print('  python3 brutal_docx_repair.py "MyBook.docx"')
    sys.exit(1)

infile = Path(sys.argv[1])

if not infile.exists():
    raise SystemExit(f"File not found: {infile}")

outfile = infile.with_name(infile.stem + "_BRUTAL_REPAIR.docx")
reportfile = infile.with_name(infile.stem + "_BRUTAL_REPAIR_report.txt")

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", NS_W)

def localname(tag):
    return tag.split("}", 1)[-1] if "}" in tag else tag

# These are removed completely.
# tbl = tables
# drawing/pict/object = images, text boxes, floating objects, embedded objects
REMOVE_WHOLE = {
    "tbl",
    "drawing",
    "pict",
    "object",

    # Comments/bookmark markers
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "bookmarkStart",
    "bookmarkEnd",

    # Word field machinery
    "fldChar",
    "instrText",

    # Tracked-change deletions / moved-away text
    "del",
    "moveFrom",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
}

# These wrappers are removed, but their visible text/content is kept.
UNWRAP = {
    "hyperlink",
    "ins",
    "moveTo",
    "smartTag",
    "sdt",
}

COUNT_INTERESTING = {
    "tbl": "tables",
    "drawing": "drawings/images/text boxes",
    "pict": "old Word pictures",
    "object": "embedded objects",
    "commentRangeStart": "comment starts",
    "commentReference": "comment refs",
    "bookmarkStart": "bookmarks",
    "fldChar": "field chars",
    "instrText": "field instructions",
    "hyperlink": "hyperlinks",
    "footnoteReference": "footnote refs",
    "endnoteReference": "endnote refs",
    "sectPr": "section breaks/properties",
    "br": "breaks",
    "del": "tracked deletions",
    "ins": "tracked insertions",
    "moveFrom": "tracked move-from",
    "moveTo": "tracked move-to",
    "sdt": "content controls",
}

def count_elements(root):
    c = Counter()
    for elem in root.iter():
        name = localname(elem.tag)
        if name in COUNT_INTERESTING:
            c[COUNT_INTERESTING[name]] += 1
    return c

def unwrap_child(parent, child, index):
    """
    Remove child but keep its children in the same position.
    Special handling for w:sdt: keep only sdtContent children if present.
    """
    name = localname(child.tag)

    if name == "sdt":
        content = None
        for sub in child:
            if localname(sub.tag) == "sdtContent":
                content = sub
                break
        new_children = list(content) if content is not None else list(child)
    else:
        new_children = list(child)

    parent.remove(child)

    for offset, grandchild in enumerate(new_children):
        parent.insert(index + offset, grandchild)

    return len(new_children)

def clean_tree(elem, removed_counter):
    i = 0

    while i < len(elem):
        child = elem[i]
        name = localname(child.tag)

        if name in REMOVE_WHOLE:
            removed_counter[name] += 1
            elem.remove(child)
            continue

        if name in UNWRAP:
            removed_counter["unwrapped_" + name] += 1
            unwrap_child(elem, child, i)
            continue

        clean_tree(child, removed_counter)
        i += 1

def is_word_xml_part(name):
    """
    Clean the main document and related visible-text parts.
    Avoid styles/settings/theme because stripping those can damage formatting globally.
    """
    if not name.startswith("word/") or not name.endswith(".xml"):
        return False

    if name == "word/document.xml":
        return True

    if name.startswith("word/header") and name.endswith(".xml"):
        return True

    if name.startswith("word/footer") and name.endswith(".xml"):
        return True

    if name in {
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    }:
        return True

    return False

with zipfile.ZipFile(infile, "r") as zin:
    files = {name: zin.read(name) for name in zin.namelist()}

before_total = Counter()
after_total = Counter()
removed_total = Counter()
processed_parts = []

new_files = dict(files)

for name, data in files.items():
    if not is_word_xml_part(name):
        continue

    try:
        root = ET.fromstring(data)
    except ET.ParseError as e:
        print(f"Could not parse {name}: {e}")
        continue

    before = count_elements(root)
    removed = Counter()

    clean_tree(root, removed)

    after = count_elements(root)

    before_total.update(before)
    after_total.update(after)
    removed_total.update(removed)
    processed_parts.append(name)

    new_files[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

media_files = [n for n in files if n.startswith("word/media/")]
embedding_files = [n for n in files if n.startswith("word/embeddings/")]

with zipfile.ZipFile(outfile, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, data in new_files.items():
        zout.writestr(name, data)

lines = []
lines.append(f"Input:  {infile}")
lines.append(f"Output: {outfile}")
lines.append("")
lines.append("Processed XML parts:")
for p in processed_parts:
    lines.append(f"  - {p}")

lines.append("")
lines.append("Package files:")
lines.append(f"  media files:      {len(media_files)}")
lines.append(f"  embedding files:  {len(embedding_files)}")

lines.append("")
lines.append("COUNTS BEFORE:")
for k, v in sorted(before_total.items()):
    lines.append(f"  {k}: {v}")

lines.append("")
lines.append("REMOVED / UNWRAPPED:")
for k, v in sorted(removed_total.items()):
    lines.append(f"  {k}: {v}")

lines.append("")
lines.append("COUNTS AFTER:")
for k, v in sorted(after_total.items()):
    lines.append(f"  {k}: {v}")

reportfile.write_text("\n".join(lines), encoding="utf-8")

print("Done.")
print(f"Wrote repaired file: {outfile}")
print(f"Wrote report:        {reportfile}")
print("")
print("Before counts:")
for k, v in sorted(before_total.items()):
    print(f"  {k}: {v}")

print("")
print("Removed/unwrapped:")
for k, v in sorted(removed_total.items()):
    print(f"  {k}: {v}")
