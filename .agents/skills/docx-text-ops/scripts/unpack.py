"""Unpack a Word (.docx) file into an editable XML directory.

Extracts the ZIP archive and pretty-prints all XML files in-place.

Usage:
    python unpack.py <file.docx> <output_dir>

Example:
    python unpack.py document.docx unpacked/
"""

import argparse
import re
import sys
import zipfile
import xml.dom.minidom
from pathlib import Path

SMART_QUOTE_MAP = {
    "“": "&#x201C;",
    "”": "&#x201D;",
    "‘": "&#x2018;",
    "’": "&#x2019;",
}

DOCX_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}


def unpack(input_file: str, output_dir: str) -> str:
    src = Path(input_file)
    dst = Path(output_dir)

    if not src.exists():
        return f"Error: {input_file} does not exist"
    if src.suffix.lower() not in DOCX_EXTENSIONS:
        return f"Error: {input_file} must be a .docx / .docm / .dotx / .dotm file"

    try:
        dst.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(src, "r") as zf:
            zf.extractall(dst)

        xml_files = list(dst.rglob("*.xml"))
        rels_files = list(dst.rglob("*.rels"))
        for f in xml_files:
            _pretty_print(f)
            _escape_smart_quotes(f)
        for f in rels_files:
            _escape_smart_quotes(f)

        return f"Unpacked {input_file} ({len(xml_files) + len(rels_files)} XML files)"
    except zipfile.BadZipFile:
        return f"Error: {input_file} is not a valid Word file"
    except Exception as e:
        return f"Error: {e}"


def _pretty_print(path: Path) -> None:
    try:
        raw = path.read_bytes()
        m = re.match(rb"<\?xml[^?]*\?>", raw)
        orig_decl = m.group().decode("utf-8") if m else None

        dom = xml.dom.minidom.parseString(raw)
        pretty = dom.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
        if orig_decl:
            pretty = re.sub(r"<\?xml[^?]*\?>", orig_decl, pretty, count=1)
        path.write_text(pretty, encoding="utf-8")
    except Exception:
        pass


def _escape_smart_quotes(path: Path) -> None:
    try:
        text = path.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_MAP.items():
            text = text.replace(char, entity)
        path.write_text(text, encoding="utf-8")
    except Exception:
        pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Unpack a Word file for XML editing")
    parser.add_argument("input_file", help="Word file to unpack (.docx or .docm)")
    parser.add_argument("output_dir", help="Output directory")
    args = parser.parse_args()

    msg = unpack(args.input_file, args.output_dir)
    print(msg)
    if "Error" in msg:
        sys.exit(1)
