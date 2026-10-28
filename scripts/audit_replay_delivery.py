"""Audit preservation, figures and mathematical indices in the exact delivery file."""
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile
from audit_replay_numbers import audit

RUN = Path(__file__).resolve().parents[1]/"outputs/chat-replay-20261028"
DELIVERY = RUN/"delivery/JTICE_OfficeCLI_聊天室復現完成版.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def load(path):
    with zipfile.ZipFile(path) as z:
        return ET.fromstring(z.read("word/document.xml"))


def addins(root):
    stack = []
    result = []
    for n in root.iter():
        if n.tag == W+"fldChar":
            kind = n.get(W+"fldCharType")
            if kind == "begin":
                stack.append({"code": "", "result": "", "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                f = stack.pop()
                if "ADDIN EN." in f["code"]:
                    result.append((f["code"], f["result"]))
        elif stack and n.tag == W+"instrText":
            stack[-1]["code"] += n.text or ""
        elif stack and n.tag == W+"t" and stack[-1]["separated"]:
            stack[-1]["result"] += n.text or ""
    for f in root.iter(W+"fldSimple"):
        if "ADDIN EN." in f.get(W+"instr", ""):
            result.append((f.get(W+"instr"), "".join(t.text or "" for t in f.iter(W+"t"))))
    return result


source = next((RUN/"inputs").glob("*_v6_*.docx"))
original = load(source)
actual = load(DELIVERY)
fields_source, fields_actual = addins(original), addins(actual)
assert fields_source == fields_actual and len(fields_actual) == 26
assert [ET.tostring(x, encoding="unicode").split(">", 1)[1] for x in original.iter(M+"oMath")] == [ET.tostring(x, encoding="unicode").split(">", 1)[1] for x in actual.iter(M+"oMath")]
body = actual.find(W+"body")
active = False
indices = 0
failures = []
for block in body:
    block_text = "".join(t.text or "" for t in block.iter(W+"t"))
    if block_text == "4 結果與討論":
        active = True
    if block_text == "5 結論":
        break
    if not active:
        continue
    paragraphs = [block] if block.tag == W+"p" else list(block.iter(W+"p"))
    for p in paragraphs:
        chars = []
        for r in p.iter(W+"r"):
            va = r.find(W+"rPr/"+W+"vertAlign")
            kind = va.get(W+"val") if va is not None else "baseline"
            chars.extend((char, kind) for t in r.iter(W+"t") for char in (t.text or ""))
        text = "".join(c for c, _ in chars)
        for pattern, offset in (("H2S", 1), ("SO2", 2)):
            for match in re.finditer(pattern, text):
                if match.start() >= 4 and text[match.start()-4:match.start()] == "B35_":
                    continue
                indices += 1
                if chars[match.start()+offset][1] != "subscript":
                    failures.append({"chemical": pattern, "paragraph": text[:160]})
assert not failures, failures[:8]
numbers = audit(DELIVERY, 15, RUN/"inputs/data")
figures = json.loads((RUN/"delivery/regenerated-figure-provenance.json").read_text(encoding="utf-8-sig"))
assert len(figures) == 10
with zipfile.ZipFile(DELIVERY) as z:
    image_hashes = {hashlib.sha256(z.read(n)).hexdigest().upper() for n in z.namelist() if n.startswith("word/media/")}
assert all(f["sha256"] in image_hashes for f in figures)
report = {"status": "PASS", "document": str(DELIVERY), "source": str(source),
          "genuine_endnote_fields_preserved": len(fields_actual),
          "endnote_code_and_cached_result_digest": hashlib.sha256(json.dumps(fields_actual).encode()).hexdigest(),
          "math_objects_preserved": len(list(actual.iter(M+"oMath"))),
          "chapter4_native_chemical_subscripts_checked": indices,
          "inline_symbol_failures": failures,
          "regenerated_figures_embedded_by_hash": len(figures), "numbers": numbers}
(RUN/"logs/delivery-audit.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(report, ensure_ascii=False), flush=True)
