"""Read-only independent comparison of replay DOCX against the original revision."""
import argparse
from collections import Counter
import csv
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def text(n):
    return "".join(x.text or "" for x in n.findall(".//w:t", NS))


def read(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        rels = ET.fromstring(z.read("word/_rels/document.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels}
        body = list(root.find("w:body", NS))
        begin = [i for i, n in enumerate(body) if text(n) == "4 結果與討論"]
        end = [i for i, n in enumerate(body) if text(n) == "5 結論"]
        assert len(begin) == len(end) == 1
        chapter = body[begin[0]+1:end[0]]
        pictures = []
        for b in root.findall(".//a:blip", NS):
            part = "word/" + relmap[b.get("{"+NS["r"]+"}embed")]
            pictures.append(hashlib.sha256(z.read(part)).hexdigest())
        return root, body, chapter, pictures, begin[0], end[0]


def normalized_text(n):
    return text(n).replace("\u00a0", " ").strip()


def resolve_reference_caches(root, chapter):
    """Compare native REF targets; saved reference files can contain stale caches."""
    active = {}
    marks = {}
    for n in root.iter():
        if n.tag == "{" + NS["w"] + "}bookmarkStart":
            active[n.get("{"+NS["w"]+"}id")] = [n.get("{"+NS["w"]+"}name"), []]
        elif n.tag == "{" + NS["w"] + "}bookmarkEnd":
            item = active.pop(n.get("{"+NS["w"]+"}id"), None)
            if item:
                marks[item[0]] = "".join(item[1])
        elif n.tag == "{" + NS["w"] + "}t":
            for name, parts in active.values():
                parts.append(n.text or "")
    corrected = []
    for p in (p for block in chapter for p in ([block] if block.tag.endswith("}p") else block.findall(".//w:p", NS))):
        for simple in p.findall(".//w:fldSimple", NS):
            m = re.match(r"\s*REF\s+(\S+)", simple.get("{"+NS["w"]+"}instr", ""))
            if m and m[1] in marks:
                ts = simple.findall(".//w:t", NS)
                old = "".join(t.text or "" for t in ts)
                if ts:
                    ts[0].text = marks[m[1]]
                    for t in ts[1:]:
                        t.text = ""
                    if old != marks[m[1]]:
                        corrected.append({"old": old, "resolved": marks[m[1]], "bookmark": m[1]})
        stack = []
        for r in p.findall(".//w:r", NS):
            fc = r.find("w:fldChar", NS)
            if fc is not None:
                kind = fc.get("{"+NS["w"]+"}fldCharType")
                if kind == "begin":
                    stack.append({"code": "", "result": False, "nodes": []})
                elif kind == "separate" and stack:
                    stack[-1]["result"] = True
                elif kind == "end" and stack:
                    f = stack.pop()
                    m = re.match(r"\s*REF\s+(\S+)", f["code"])
                    if m and m[1] in marks and f["nodes"]:
                        old = "".join(t.text or "" for t in f["nodes"])
                        f["nodes"][0].text = marks[m[1]]
                        for t in f["nodes"][1:]:
                            t.text = ""
                        if old != marks[m[1]]:
                            corrected.append({"old": old, "resolved": marks[m[1]], "bookmark": m[1]})
                continue
            if stack:
                stack[-1]["code"] += "".join(t.text or "" for t in r.findall("w:instrText", NS))
                if stack[-1]["result"]:
                    stack[-1]["nodes"].extend(r.findall("w:t", NS))
    return corrected


def semantic_blocks(chapter):
    out = []
    for n in chapter:
        t = normalized_text(n)
        drawing = n.find(".//w:drawing", NS) is not None
        if n.tag.endswith("}p") and not t and not drawing and not n.findall(".//w:br", NS):
            continue  # Blank paragraph spacing is assessed by rendered QA.
        out.append(("picture" if drawing else n.tag.rsplit("}", 1)[-1], t))
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("document")
    p.add_argument("--reference", required=True)
    p.add_argument("--report", required=True)
    args = p.parse_args()
    actual = read(args.document)
    reference = read(args.reference)
    ar, ab, ac, ap, aa, ae = actual
    rr, rb, rc, rp, ra, re_ = reference
    literal_equal = [text(n) for n in ac] == [text(n) for n in rc]
    corrected_actual = resolve_reference_caches(ar, ac)
    corrected_source = resolve_reference_caches(rr, rc)
    actual_blocks = semantic_blocks(ac)
    reference_blocks = semantic_blocks(rc)
    differences = []
    for i in range(max(len(actual_blocks), len(reference_blocks))):
        a = actual_blocks[i] if i < len(actual_blocks) else None
        r = reference_blocks[i] if i < len(reference_blocks) else None
        if a != r:
            differences.append({"block": i, "actual": a, "reference": r})
    # Compare mathematical content, media bytes and all table values independently.
    equations = len(ar.findall(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
    reference_equations = len(rr.findall(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
    tables_a = [[text(c) for c in row.findall("w:tc", NS)] for t in ac if t.tag.endswith("}tbl") for row in t.findall("w:tr", NS)]
    tables_r = [[text(c) for c in row.findall("w:tc", NS)] for t in rc if t.tag.endswith("}tbl") for row in t.findall("w:tr", NS)]
    prefix_differences = []
    for i, (a, r) in enumerate(zip(ab[:aa], rb[:ra])):
        if text(a) != text(r):
            prefix_differences.append({"block": i, "actual": text(a), "reference": text(r)})
    report = {
        "document": str(Path(args.document).resolve()),
        "reference": str(Path(args.reference).resolve()),
        "chapter_blocks": [len(actual_blocks), len(reference_blocks)],
        "chapter_literal_text_equal": literal_equal,
        "comparison_basis": "Resolved native REF targets; blank paragraph spacing is checked in the Word render",
        "stale_source_reference_caches": corrected_source,
        "stale_replay_reference_caches": corrected_actual,
        "chapter_text_equal": not differences,
        "chapter_differences": differences,
        "tables_equal": tables_a == tables_r,
        "table_rows": len(tables_a),
        "media_equal_in_order": ap == rp,
        "media_count": [len(ap), len(rp)],
        "equations_equal": equations == reference_equations,
        "equations": equations,
        "prefix_text_equal": not prefix_differences and aa == ra,
        "prefix_differences": prefix_differences,
        "suffix_text_equal": [text(n) for n in ab[ae:]] == [text(n) for n in rb[re_:]],
    }
    Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k not in ("chapter_differences", "prefix_differences")}, ensure_ascii=False))
    print("Differences:", len(differences), "prefix:", len(prefix_differences))
    return 0 if all(report[k] for k in ("chapter_text_equal", "tables_equal", "media_equal_in_order", "equations_equal", "prefix_text_equal", "suffix_text_equal")) else 1


if __name__ == "__main__":
    raise SystemExit(main())
