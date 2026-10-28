"""Create a page inventory and contact sheets for the native Word PDF render."""
from pathlib import Path
import json
from PIL import Image, ImageDraw
import pypdfium2 as pdfium

RUN = Path(__file__).resolve().parents[1]/"outputs/chat-replay-20261028"
RENDER = RUN/"render/delivery"
pdfpath = next(RENDER.glob("*.pdf"))
pdf = pdfium.PdfDocument(str(pdfpath))
inventory = []
for i in range(len(pdf)):
    page = pdf[i]
    tp = page.get_textpage()
    text = tp.get_text_range()
    inventory.append({"page": i+1, "characters": len(text.strip()),
                      "start": text[:120], "tables": [k for k in ("Table 9a", "Table 9b", "Table 10a", "Table 10b", "Table 12", "Table 13") if k in text],
                      "existing_broken_reference": "找不到參照來源" in text,
                      "chapter4_heading": "4 結果與討論" in text,
                      "literal_escape_leak": "\\v" in text or "@@REF" in text or "@@CAP" in text})
    tp.close()
    page.close()
pdf.close()
assert not any(p["literal_escape_leak"] for p in inventory)
for start in range(0, len(inventory), 12):
    thumbw, thumbh = 390, 552
    sheet = Image.new("RGB", (thumbw*3, (thumbh+28)*4), "#d8d8d8")
    draw = ImageDraw.Draw(sheet)
    for k, entry in enumerate(inventory[start:start+12]):
        with Image.open(RENDER/f'page-{entry["page"]}.png') as im:
            im.thumbnail((thumbw, thumbh))
            x, y = (k%3)*thumbw, (k//3)*(thumbh+28)
            draw.text((x+8, y+7), f'Page {entry["page"]}', fill="black")
            sheet.paste(im, (x+(thumbw-im.width)//2, y+28))
    sheet.save(RENDER/f"contact-{start//12+1}.png")
(RUN/"logs/render-inventory.json").write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding="utf-8")
print("Pages:", len(inventory))
print("Chapter 4:", [p["page"] for p in inventory if p["chapter4_heading"]])
print("Table pages:", [(p["page"], p["tables"]) for p in inventory if p["tables"]])
print("Very short text pages (may contain figures):", [p["page"] for p in inventory if p["characters"] < 100])
print("Known original reference error pages:", [p["page"] for p in inventory if p["existing_broken_reference"]])
