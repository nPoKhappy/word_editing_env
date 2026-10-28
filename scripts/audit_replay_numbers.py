"""Audit actual Word table values and formatting against independent raw CSVs."""
import argparse
import csv
from decimal import Decimal, ROUND_HALF_UP
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIVE = ["No-Gain"] + [f"Difficult-GradNorm ratio={r}" for r in ("0.005", "0.01", "0.015", "0.02")]
FOUR = ["No-Gain", "Gain", "Gain + SS=0.01", "Difficult-GradNorm ratio=0.01"]


def text(n):
    return "".join(x.text or "" for x in n.iter(W+"t"))


def scaled(value):
    v = value * 1000
    return format(v.quantize(Decimal(1).scaleb(v.adjusted()-4)), "f") if v else "0.0000"


def percent(v, base, decimals):
    return f'{((v-base)/base*100).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP):+.{decimals}f}%'


def run_format(cell, token, styles):
    matches = [r for r in cell.iter(W+"r") if token in text(r)]
    assert matches, (token, text(cell))
    r = matches[0]
    b = r.find(W+"rPr/"+W+"b")
    ps = cell.find(W+"p/"+W+"pPr/"+W+"pStyle")
    style = ps.get(W+"val") if ps is not None else "Normal"
    inherited = False
    seen = set()
    while style and style not in seen:
        seen.add(style)
        s = styles.get(style)
        if s is None:
            break
        sb = s.find(W+"rPr/"+W+"b")
        if sb is not None:
            inherited = sb.get(W+"val", "1") not in ("0", "false", "off")
            break
        based = s.find(W+"basedOn")
        style = based.get(W+"val") if based is not None else None
    bold = (b.get(W+"val", "1") not in ("0", "false", "off")) if b is not None else inherited
    color = r.find(W+"rPr/"+W+"color")
    return bold, color.get(W+"val", "000000").upper() if color is not None else "000000"


def audit(path, version, data):
    lookup = {}
    for filename in ("seven_model_validation_rmse_values.csv", "seven_model_test_rmse_values.csv"):
        with (data/filename).open(encoding="utf-8-sig", newline="") as fp:
            for row in csv.DictReader(fp):
                key = (row["model"], row["case"], row["target"], row["evaluation"])
                assert key not in lookup
                lookup[key] = Decimal(row["RMSE"])
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        styles = {s.get(W+"styleId"): s for s in ET.fromstring(z.read("word/styles.xml")) if s.tag == W+"style"}
    body = list(root.find(W+"body"))
    active = False
    title = ""
    report = {"version": version, "raw_rmse_cells": 0, "percentage_cells": 0,
              "minimum_bold_cells": 0, "baseline_color_cells": 0, "selected_rho_gray_cells": 0,
              "table_rows": 0, "tables": 0}
    for node in body:
        if text(node) == "4 結果與討論":
            active = True
        if text(node) == "5 結論":
            break
        if not active:
            continue
        if node.tag == W+"p" and text(node).startswith("Table "):
            title = text(node)
        if node.tag != W+"tbl":
            continue
        rows = node.findall(W+"tr")
        first = [text(c) for c in rows[1].findall(W+"tc")]
        if not first[0].startswith("R5"):
            continue
        report["tables"] += 1
        split = "相對變化率" in title
        validation = "驗證" in title
        target_for_split = "B35_H2S" if "H2S" in title else "B35_SO2"
        models = FIVE if validation or version < 15 else FOUR
        for row in rows[1:]:
            cells = row.findall(W+"tc")
            values = [text(c) for c in cells]
            case = values[0]
            if split:
                target, metric = target_for_split, values[1]
            else:
                chemical, metric = values[1].split()
                target = "B35_"+chemical
            exact = [lookup[(m, case, target, metric)] for m in models]
            minimum = min(exact)
            columns = list(range(1, 5)) if split else list(range(len(models)))
            for j in columns:
                cell = cells[j+1 if split else j+2]
                raw = values[j+1 if split else j+2]
                if split:
                    want = percent(exact[j], exact[0], 1)
                    assert raw == want, (version, case, target, metric, raw, want)
                    report["percentage_cells"] += 1
                else:
                    want = scaled(exact[j])
                    assert raw.startswith(want), (version, case, target, metric, raw, want)
                    report["raw_rmse_cells"] += 1
                    if version == 12 and validation and j:
                        assert raw == want+percent(exact[j], exact[0], 2)
                        report["percentage_cells"] += 1
                bold, color = run_format(cell, want, styles)
                assert bold == (exact[j] == minimum), (version, case, target, metric, j, "bold")
                report["minimum_bold_cells"] += 1
                if validation and version >= 11 and j:
                    expected = "008000" if exact[j] < exact[0] else "C00000" if exact[j] > exact[0] else "000000"
                    assert color == expected, (version, case, target, metric, j, color, expected)
                    report["baseline_color_cells"] += 1
                if split and j == 2:
                    fill = cell.find(W+"tcPr/"+W+"shd")
                    assert fill is not None and fill.get(W+"fill").upper() == "EEEEEE"
                    report["selected_rho_gray_cells"] += 1
            report["table_rows"] += 1
    assert report["table_rows"] == 102, report
    if version >= 12:
        assert report["percentage_cells"] == 120, report
    report["status"] = "PASS"
    return report


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--run", type=Path, default=Path("outputs/chat-replay-20261028"))
    args = p.parse_args()
    reports = []
    for v in range(9, 16):
        file = args.run/"stages"/("v15-officecli-03.docx" if v == 15 else "v12-officecli-02.docx" if v == 12 else f"v{v}-officecli.docx")
        try:
            result = audit(file, v, args.run/"inputs/data")
        except Exception as exc:
            result = {"version": v, "status": "FAIL", "error": repr(exc)}
        reports.append(result)
        print(json.dumps(result), flush=True)
    (args.run/"logs/numeric-audit.json").write_text(json.dumps(reports, indent=2), encoding="utf-8")
    raise SystemExit(0 if all(r["status"] == "PASS" for r in reports) else 1)
