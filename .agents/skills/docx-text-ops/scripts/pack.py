"""Repack an unpacked Word directory back into a .docx file.

Places [Content_Types].xml first (uncompressed) per the OPC spec,
then _rels/, then all remaining files with deflate compression.

Usage:
    python pack.py <unpacked_dir> <output.docx>

Example:
    python pack.py unpacked/ edited.docx
"""

import argparse
import sys
import zipfile
from pathlib import Path

# OPC spec requires [Content_Types].xml to be stored without compression
_UNCOMPRESSED = {"[Content_Types].xml"}

DOCX_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}


def _sort_key(rel: str) -> tuple:
    if rel == "[Content_Types].xml":
        return (0, rel)
    if rel.startswith("_rels/"):
        return (1, rel)
    return (2, rel)


def pack(input_dir: str, output_file: str) -> str:
    src = Path(input_dir)
    dst = Path(output_file)

    if not src.is_dir():
        return f"Error: {input_dir} is not a directory"
    if dst.suffix.lower() not in DOCX_EXTENSIONS:
        return f"Error: {output_file} must be a .docx / .docm / .dotx / .dotm file"

    try:
        files = sorted(
            (f for f in src.rglob("*") if f.is_file()),
            key=lambda f: _sort_key(f.relative_to(src).as_posix()),
        )

        with zipfile.ZipFile(dst, "w") as zf:
            for f in files:
                arcname = f.relative_to(src).as_posix()
                compress = (
                    zipfile.ZIP_STORED
                    if arcname in _UNCOMPRESSED
                    else zipfile.ZIP_DEFLATED
                )
                zf.write(f, arcname, compress_type=compress)

        return f"Packed {len(files)} files into {output_file}"
    except Exception as e:
        return f"Error: {e}"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Repack an unpacked Word directory")
    parser.add_argument("input_dir", help="Unpacked directory")
    parser.add_argument("output_file", help="Output .docx file")
    args = parser.parse_args()

    msg = pack(args.input_dir, args.output_file)
    print(msg)
    if "Error" in msg:
        sys.exit(1)
