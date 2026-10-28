"""Select an unambiguous fragment for every newly built table, then run native QA."""
import argparse
import json
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text(n):
    return "".join(t.text or "" for t in n.iter(W+"t"))


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("document")
    p.add_argument("--report", required=True)
    args = p.parse_args()
    with zipfile.ZipFile(args.document) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    tables = list(root.iter(W+"tbl"))
    texts = [text(t) for t in tables]
    active = False
    needles = []
    for n in root.find(W+"body"):
        if text(n) == "4 結果與討論":
            active = True
        if text(n) == "5 結論":
            break
        if active and n.tag == W+"tbl":
            candidates = [text(t) for t in n.iter(W+"tc") if text(t) and not list(t.iter(W+"br"))]
            found = [t for t in candidates if sum(t in s for s in texts) == 1]
            assert found, "No unique cell fragment; construct an explicit native-range audit"
            needles.append(max(found, key=len))
    cmd = [sys.executable, "-X", "utf8", str(Path(__file__).with_name("verify_word_font_sizes.py")),
           args.document, "--directories", "--allow-no-directories", "--timeout", "240", "--report", args.report]
    for needle in needles:
        cmd.append("--table-text="+needle)
    Path(args.report+".targets.json").write_text(json.dumps(needles, ensure_ascii=False, indent=2), encoding="utf-8")
    return subprocess.run(cmd, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
