#!/usr/bin/env python3
"""
inspect_design.py — Read-only design audit for .docx files.

Reports: THEME FONTS, THEME COLORS, STYLES, INLINE OVERRIDES

Usage:
    python3 skills/docx-design-ops/scripts/inspect_design.py document.docx
    python3 skills/docx-design-ops/scripts/inspect_design.py document.docx --json
    python3 skills/docx-design-ops/scripts/inspect_design.py document.docx --unpacked unpacked/
"""

import argparse
import json
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def W(tag):
    return f"{{{NS_W}}}{tag}"


def A(tag):
    return f"{{{NS_A}}}{tag}"


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def parse_theme(xml_bytes):
    root = ET.fromstring(xml_bytes)
    result = {"majorFont": None, "minorFont": None, "colors": {}}

    font_scheme = root.find(f".//{A('fontScheme')}")
    if font_scheme is not None:
        major = font_scheme.find(f"{A('majorFont')}/{A('latin')}")
        if major is not None:
            result["majorFont"] = major.get("typeface")
        minor = font_scheme.find(f"{A('minorFont')}/{A('latin')}")
        if minor is not None:
            result["minorFont"] = minor.get("typeface")

    clr_scheme = root.find(f".//{A('clrScheme')}")
    if clr_scheme is not None:
        for child in clr_scheme:
            slot = child.tag.split("}")[-1]
            srgb = child.find(A("srgbClr"))
            sysclr = child.find(A("sysClr"))
            if srgb is not None:
                result["colors"][slot] = srgb.get("val", "").upper()
            elif sysclr is not None:
                result["colors"][slot] = sysclr.get("lastClr", "").upper()

    return result


def _extract_rpr(rpr):
    if rpr is None:
        return {}
    out = {}

    fonts = rpr.find(W("rFonts"))
    if fonts is not None:
        f = {}
        for attr in ("ascii", "hAnsi", "eastAsia", "cs",
                     "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            v = fonts.get(W(attr))
            if v:
                f[attr] = v
        if f:
            out["rFonts"] = f

    sz = rpr.find(W("sz"))
    if sz is not None:
        raw = sz.get(W("val"))
        if raw:
            out["sz"] = int(raw)

    color = rpr.find(W("color"))
    if color is not None:
        v = color.get(W("val"))
        if v:
            out["color"] = v
        tc = color.get(W("themeColor"))
        if tc:
            out["themeColor"] = tc

    out["bold"] = rpr.find(W("b")) is not None
    out["italic"] = rpr.find(W("i")) is not None
    return out


def _extract_ppr(ppr):
    if ppr is None:
        return {}
    out = {}

    spacing = ppr.find(W("spacing"))
    if spacing is not None:
        s = {}
        for attr in ("before", "after", "line", "lineRule"):
            v = spacing.get(W(attr))
            if v:
                s[attr] = v
        if s:
            out["spacing"] = s

    jc = ppr.find(W("jc"))
    if jc is not None:
        v = jc.get(W("val"))
        if v:
            out["jc"] = v

    return out


def parse_styles(xml_bytes):
    root = ET.fromstring(xml_bytes)
    styles = []

    doc_defaults = root.find(W("docDefaults"))
    if doc_defaults is not None:
        rpr_d = doc_defaults.find(f"{W('rPrDefault')}/{W('rPr')}")
        ppr_d = doc_defaults.find(f"{W('pPrDefault')}/{W('pPr')}")
        styles.append({
            "styleId": "__docDefaults__",
            "type": "defaults",
            "name": "Document Defaults",
            "basedOn": None,
            "rPr": _extract_rpr(rpr_d),
            "pPr": _extract_ppr(ppr_d),
        })

    for elem in root.findall(W("style")):
        sid = elem.get(W("styleId"))
        stype = elem.get(W("type"))
        name_e = elem.find(W("name"))
        based_e = elem.find(W("basedOn"))
        styles.append({
            "styleId": sid,
            "type": stype,
            "name": name_e.get(W("val")) if name_e is not None else sid,
            "basedOn": based_e.get(W("val")) if based_e is not None else None,
            "rPr": _extract_rpr(elem.find(W("rPr"))),
            "pPr": _extract_ppr(elem.find(W("pPr"))),
        })

    return styles


def count_inline_overrides(xml_bytes):
    root = ET.fromstring(xml_bytes)
    para_count = 0
    run_count = 0
    examples = []

    for para in root.iter(W("p")):
        ppr = para.find(W("pPr"))
        if ppr is not None:
            non_structural = [
                c for c in ppr
                if c.tag not in (W("pStyle"), W("numPr"), W("sectPr"),
                                 W("jc"), W("tabs"), W("suppressAutoHyphens"))
                and c.find(W("spacing")) is None and c.tag != W("spacing")
            ]
            spacing_override = ppr.find(W("spacing"))
            if non_structural or spacing_override is not None:
                para_count += 1

        for run in para.findall(W("r")):
            rpr = run.find(W("rPr"))
            if rpr is None:
                continue
            has_override = any(
                rpr.find(W(t)) is not None
                for t in ("rFonts", "color", "sz", "b", "i")
            )
            if has_override:
                run_count += 1
                if len(examples) < 3:
                    t_e = run.find(W("t"))
                    if t_e is not None and t_e.text:
                        examples.append(t_e.text[:50])

    return para_count, run_count, examples


# ---------------------------------------------------------------------------
# Main inspector
# ---------------------------------------------------------------------------

def inspect(docx_path, unpacked_dir=None):
    if unpacked_dir:
        base = Path(unpacked_dir)

        def read(member):
            p = base / member
            return p.read_bytes() if p.exists() else None
    else:
        zf = zipfile.ZipFile(docx_path)

        def read(member):
            try:
                return zf.read(member)
            except KeyError:
                return None

    result = {"theme": {}, "styles": [], "inlineOverrides": {}}

    theme_bytes = read("word/theme/theme1.xml")
    if theme_bytes:
        result["theme"] = parse_theme(theme_bytes)

    styles_bytes = read("word/styles.xml")
    if styles_bytes:
        result["styles"] = parse_styles(styles_bytes)

    doc_bytes = read("word/document.xml")
    if doc_bytes:
        p_count, r_count, examples = count_inline_overrides(doc_bytes)
        result["inlineOverrides"] = {
            "paragraphsWithDirectPpr": p_count,
            "runsWithDirectRpr": r_count,
            "examples": examples,
        }

    if not unpacked_dir:
        zf.close()

    return result


# ---------------------------------------------------------------------------
# Human-readable output
# ---------------------------------------------------------------------------

def print_human(data):
    SEP = "=" * 60

    print(SEP)
    print("THEME FONTS")
    print(SEP)
    t = data.get("theme", {})
    print(f"  Major font (headings): {t.get('majorFont') or '(not found)'}")
    print(f"  Minor font (body):     {t.get('minorFont') or '(not found)'}")
    print()

    print(SEP)
    print("THEME COLORS")
    print(SEP)
    for slot, val in t.get("colors", {}).items():
        print(f"  {slot:<12}  #{val}")
    print()

    print(SEP)
    print("STYLES  (paragraph / character / defaults)")
    print(SEP)
    for s in data.get("styles", []):
        if s["type"] not in ("paragraph", "character", "defaults"):
            continue
        print(f"\n  [{s['styleId']}]  \"{s['name']}\"  type={s['type']}")
        if s.get("basedOn"):
            print(f"    basedOn: {s['basedOn']}")

        rpr = s.get("rPr", {})
        fonts = rpr.get("rFonts", {})
        if fonts:
            direct = {k: v for k, v in fonts.items()
                      if not k.endswith("Theme") and not k.endswith("theme")}
            themed = {k: v for k, v in fonts.items()
                      if k.endswith("Theme") or k.endswith("theme")}
            if direct:
                print(f"    font (direct): {direct}")
            if themed:
                print(f"    font (theme):  {themed}")
        if "sz" in rpr:
            pt = rpr["sz"] / 2
            print(f"    size:    {rpr['sz']} half-pts  ({pt:.1f}pt)")
        color = rpr.get("color")
        if color:
            tc = rpr.get("themeColor")
            suffix = f"  (themeColor: {tc})" if tc else ""
            print(f"    color:   #{color}{suffix}")
        flags = []
        if rpr.get("bold"):
            flags.append("bold")
        if rpr.get("italic"):
            flags.append("italic")
        if flags:
            print(f"    flags:   {', '.join(flags)}")

        ppr = s.get("pPr", {})
        sp = ppr.get("spacing", {})
        if sp:
            print(f"    spacing: {sp}")
        if ppr.get("jc"):
            print(f"    align:   {ppr['jc']}")
    print()

    print(SEP)
    print("INLINE OVERRIDES  (in word/document.xml)")
    print(SEP)
    ov = data.get("inlineOverrides", {})
    print(f"  Paragraphs with direct pPr overrides: {ov.get('paragraphsWithDirectPpr', 0)}")
    print(f"  Runs with direct rPr overrides:       {ov.get('runsWithDirectRpr', 0)}")
    ex = ov.get("examples", [])
    if ex:
        print("  Examples:")
        for e in ex:
            print(f"    \"{e}\"")
    if ov.get("runsWithDirectRpr", 0) > 0:
        print()
        print("  NOTE: Inline overrides take precedence over style definitions.")
        print("  Style-level changes may not affect these runs/paragraphs.")
        print("  See design-recipes.md Recipe 5 to audit and remove them.")
    print()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Read-only design audit for .docx files",
    )
    parser.add_argument("docx", help=".docx file path")
    parser.add_argument("--json", action="store_true", help="Output JSON")
    parser.add_argument(
        "--unpacked", metavar="DIR",
        help="Use an already-unpacked directory instead of the ZIP",
    )
    args = parser.parse_args()

    path = Path(args.docx)
    if not path.exists():
        print(f"Error: not found: {path}", file=sys.stderr)
        sys.exit(1)

    data = inspect(path, unpacked_dir=args.unpacked)

    if args.json:
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        print_human(data)


if __name__ == "__main__":
    main()
