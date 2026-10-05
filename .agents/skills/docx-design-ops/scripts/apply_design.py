#!/usr/bin/env python3
"""
apply_design.py - Batch-apply a design spec to an unpacked .docx directory.

Modifies: word/theme/theme1.xml, word/styles.xml only.
Never touches: word/document.xml, any .rels file.

Usage:
    python3 skills/docx-design-ops/scripts/apply_design.py unpacked/ spec.json
    python3 skills/docx-design-ops/scripts/apply_design.py unpacked/ spec.json --dry-run

spec.json format:
{
  "theme": {
    "majorFont": "Calibri Light",
    "minorFont": "Calibri",
    "colors": {
      "accent1": "2E74B5",
      "dk1": "000000"
    }
  },
  "styles": {
    "Normal": {
      "rPr": { "rFonts": "Calibri", "sz": 22 },
      "pPr": { "spacingBefore": 0, "spacingAfter": 160, "line": 276, "lineRule": "auto" }
    },
    "Heading1": {
      "rPr": { "rFonts": "Calibri Light", "sz": 32, "color": "2E74B5", "bold": true }
    }
  }
}

rPr keys:
  rFonts      font name string (sets w:ascii, w:hAnsi, w:eastAsia, w:cs)
  sz          font size in half-points (24 = 12pt); also sets szCs
  color       6-digit hex string (no #); removes themeColor reference
  bold        true/false
  italic      true/false

pPr keys:
  spacingBefore   space before paragraph in twips (240 = 12pt)
  spacingAfter    space after paragraph in twips (160 = 8pt)
  line            line height in twips (276 = 1.15× at 12pt)
  lineRule        "auto" | "exact" | "atLeast"
"""

import argparse
import json
import shutil
import sys
from pathlib import Path
import xml.etree.ElementTree as ET

# Register namespaces so ET preserves prefixes when serialising
_NS = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o":   "urn:schemas-microsoft-com:office:office",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m":   "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v":   "urn:schemas-microsoft-com:vml",
    "wp14":"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16se":"http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
for _prefix, _uri in _NS.items():
    ET.register_namespace(_prefix, _uri)

NS_W = _NS["w"]
NS_A = _NS["a"]


def W(tag):
    return f"{{{NS_W}}}{tag}"


def A(tag):
    return f"{{{NS_A}}}{tag}"


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _get_or_create(parent, tag):
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    return child


def _remove_attrs(elem, *attr_tags):
    for t in attr_tags:
        if t in elem.attrib:
            del elem.attrib[t]


# ---------------------------------------------------------------------------
# Theme modifications
# ---------------------------------------------------------------------------

def apply_theme(root, spec, changes):
    font_scheme = root.find(f".//{A('fontScheme')}")
    if font_scheme is None and (spec.get("majorFont") or spec.get("minorFont")):
        changes.append("  WARNING: a:fontScheme not found in theme - font changes skipped")
        return

    if "majorFont" in spec and font_scheme is not None:
        latin = font_scheme.find(f"{A('majorFont')}/{A('latin')}")
        if latin is not None:
            old = latin.get("typeface", "")
            new = spec["majorFont"]
            if old != new:
                changes.append(f"  theme.majorFont: {old!r} -> {new!r}")
                latin.set("typeface", new)

    if "minorFont" in spec and font_scheme is not None:
        latin = font_scheme.find(f"{A('minorFont')}/{A('latin')}")
        if latin is not None:
            old = latin.get("typeface", "")
            new = spec["minorFont"]
            if old != new:
                changes.append(f"  theme.minorFont: {old!r} -> {new!r}")
                latin.set("typeface", new)

    if "colors" in spec:
        clr_scheme = root.find(f".//{A('clrScheme')}")
        if clr_scheme is None:
            changes.append("  WARNING: a:clrScheme not found - color changes skipped")
            return
        for slot, new_hex in spec["colors"].items():
            new_hex = new_hex.upper().lstrip("#")
            slot_elem = clr_scheme.find(A(slot))
            if slot_elem is None:
                changes.append(f"  WARNING: color slot {slot!r} not found - skipped")
                continue
            srgb = slot_elem.find(A("srgbClr"))
            sysclr = slot_elem.find(A("sysClr"))
            if srgb is not None:
                old = srgb.get("val", "")
                if old.upper() != new_hex:
                    changes.append(f"  theme.colors.{slot}: #{old} -> #{new_hex}")
                    srgb.set("val", new_hex)
            elif sysclr is not None:
                old = sysclr.get("lastClr", "")
                if old.upper() != new_hex:
                    changes.append(f"  theme.colors.{slot} (sysClr lastClr): #{old} -> #{new_hex}")
                    sysclr.set("lastClr", new_hex)
            else:
                # Create srgbClr child
                changes.append(f"  theme.colors.{slot}: (none) -> #{new_hex}")
                ET.SubElement(slot_elem, A("srgbClr")).set("val", new_hex)


# ---------------------------------------------------------------------------
# Style modifications
# ---------------------------------------------------------------------------

def _apply_rpr(style_elem, rpr_spec, style_id, changes):
    rpr = _get_or_create(style_elem, W("rPr"))

    if "rFonts" in rpr_spec:
        fname = rpr_spec["rFonts"]
        fonts = _get_or_create(rpr, W("rFonts"))
        # Remove theme references; we are setting explicit names.
        _remove_attrs(fonts, W("asciiTheme"), W("hAnsiTheme"),
                      W("eastAsiaTheme"), W("cstheme"))
        old = fonts.get(W("ascii"), "")
        if old != fname:
            changes.append(f"  style[{style_id}].rPr.rFonts: {old!r} -> {fname!r}")
        fonts.set(W("ascii"), fname)
        fonts.set(W("hAnsi"), fname)
        fonts.set(W("eastAsia"), fname)
        fonts.set(W("cs"), fname)

    if "sz" in rpr_spec:
        sz_val = str(rpr_spec["sz"])
        sz_elem = _get_or_create(rpr, W("sz"))
        old = sz_elem.get(W("val"), "")
        if old != sz_val:
            changes.append(f"  style[{style_id}].rPr.sz: {old!r} -> {sz_val!r}")
        sz_elem.set(W("val"), sz_val)
        # Always sync complex-script size
        szcs = _get_or_create(rpr, W("szCs"))
        szcs.set(W("val"), sz_val)

    if "color" in rpr_spec:
        hex_val = rpr_spec["color"].upper().lstrip("#")
        color_elem = _get_or_create(rpr, W("color"))
        old = color_elem.get(W("val"), "")
        if old.upper() != hex_val:
            changes.append(f"  style[{style_id}].rPr.color: #{old} -> #{hex_val}")
        color_elem.set(W("val"), hex_val)
        # Remove theme color references; direct hex wins.
        _remove_attrs(color_elem, W("themeColor"), W("themeTint"), W("themeShade"))

    if "bold" in rpr_spec:
        b = rpr.find(W("b"))
        if rpr_spec["bold"]:
            if b is None:
                changes.append(f"  style[{style_id}].rPr.bold: false -> true")
                ET.SubElement(rpr, W("b"))
        else:
            if b is not None:
                changes.append(f"  style[{style_id}].rPr.bold: true -> false")
                rpr.remove(b)

    if "italic" in rpr_spec:
        i = rpr.find(W("i"))
        if rpr_spec["italic"]:
            if i is None:
                changes.append(f"  style[{style_id}].rPr.italic: false -> true")
                ET.SubElement(rpr, W("i"))
        else:
            if i is not None:
                changes.append(f"  style[{style_id}].rPr.italic: true -> false")
                rpr.remove(i)


def _apply_ppr(style_elem, ppr_spec, style_id, changes):
    ppr = _get_or_create(style_elem, W("pPr"))

    spacing_keys = ("spacingBefore", "spacingAfter", "line", "lineRule")
    if any(k in ppr_spec for k in spacing_keys):
        spacing = _get_or_create(ppr, W("spacing"))
        mapping = {
            "spacingBefore": W("before"),
            "spacingAfter":  W("after"),
            "line":          W("line"),
            "lineRule":      W("lineRule"),
        }
        for spec_key, xml_attr in mapping.items():
            if spec_key in ppr_spec:
                new_val = str(ppr_spec[spec_key])
                old_val = spacing.get(xml_attr, "")
                if old_val != new_val:
                    changes.append(
                        f"  style[{style_id}].pPr.spacing.{spec_key}: {old_val!r} -> {new_val!r}"
                    )
                spacing.set(xml_attr, new_val)


def apply_styles(root, styles_spec, changes):
    for style_id, props in styles_spec.items():
        # Find style by styleId attribute
        style_elem = None
        for s in root.findall(W("style")):
            if s.get(W("styleId")) == style_id:
                style_elem = s
                break

        if style_elem is None:
            changes.append(f"  WARNING: style {style_id!r} not found - skipped")
            continue

        if "rPr" in props:
            _apply_rpr(style_elem, props["rPr"], style_id, changes)
        if "pPr" in props:
            _apply_ppr(style_elem, props["pPr"], style_id, changes)


# ---------------------------------------------------------------------------
# File I/O
# ---------------------------------------------------------------------------

def load_xml(path):
    tree = ET.parse(path)
    return tree, tree.getroot()


def save_xml(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(unpacked_dir, spec, dry_run):
    base = Path(unpacked_dir)

    theme_path = base / "word" / "theme" / "theme1.xml"
    styles_path = base / "word" / "styles.xml"

    all_changes = []

    # --- Theme ---
    if "theme" in spec:
        if not theme_path.exists():
            all_changes.append("  WARNING: word/theme/theme1.xml not found - theme changes skipped")
        else:
            _, theme_root = load_xml(theme_path)
            theme_changes = []
            apply_theme(theme_root, spec["theme"], theme_changes)
            if theme_changes:
                all_changes.append("word/theme/theme1.xml:")
                all_changes.extend(theme_changes)
                if not dry_run:
                    backup = theme_path.with_suffix(".xml.bak")
                    shutil.copy2(theme_path, backup)
                    save_xml(ET.ElementTree(theme_root), theme_path)
            else:
                all_changes.append("word/theme/theme1.xml: no changes needed")

    # --- Styles ---
    if "styles" in spec:
        if not styles_path.exists():
            all_changes.append("  WARNING: word/styles.xml not found - style changes skipped")
        else:
            _, styles_root = load_xml(styles_path)
            style_changes = []
            apply_styles(styles_root, spec["styles"], style_changes)
            if style_changes:
                all_changes.append("word/styles.xml:")
                all_changes.extend(style_changes)
                if not dry_run:
                    backup = styles_path.with_suffix(".xml.bak")
                    shutil.copy2(styles_path, backup)
                    save_xml(ET.ElementTree(styles_root), styles_path)
            else:
                all_changes.append("word/styles.xml: no changes needed")

    return all_changes


def main():
    parser = argparse.ArgumentParser(
        description="Batch-apply a design spec to an unpacked .docx directory"
    )
    parser.add_argument("unpacked_dir", help="Path to unpacked .docx directory")
    parser.add_argument("spec_file", help="Path to design spec JSON file (use - for stdin)")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Show planned changes without writing any files",
    )
    args = parser.parse_args()

    if args.spec_file == "-":
        spec = json.load(sys.stdin)
    else:
        with open(args.spec_file, encoding="utf-8") as fh:
            spec = json.load(fh)

    if not Path(args.unpacked_dir).is_dir():
        print(f"Error: directory not found: {args.unpacked_dir}", file=sys.stderr)
        sys.exit(1)

    changes = run(args.unpacked_dir, spec, dry_run=args.dry_run)

    if args.dry_run:
        print("DRY RUN - no files will be written\n")

    for line in changes:
        print(line)

    if args.dry_run:
        print("\nRun without --dry-run to apply these changes.")
    else:
        backups = list(Path(args.unpacked_dir).rglob("*.xml.bak"))
        if backups:
            print(f"\nBackups written: {[str(b) for b in backups]}")
        print("Done. Run pack.py to rebuild the .docx file.")


if __name__ == "__main__":
    main()
