"""Export DOCX with installed Microsoft Word and render its PDF pages to PNG."""

import argparse
import os
from pathlib import Path
import tempfile


def render(docx_path: Path, out_dir: Path, scale: float) -> None:
    if not docx_path.is_file():
        raise FileNotFoundError(docx_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    local_temp = Path(__file__).resolve().parents[1] / ".tmp"
    local_temp.mkdir(exist_ok=True)
    os.environ["TEMP"] = str(local_temp)
    os.environ["TMP"] = str(local_temp)
    tempfile.tempdir = str(local_temp)

    import pythoncom
    import win32com.client
    import pypdfium2 as pdfium

    pdf_path = out_dir / (docx_path.stem + ".pdf")
    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(docx_path.resolve()), ConfirmConversions=False,
            ReadOnly=True, AddToRecentFiles=False,
        )
        document.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()

    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            image = page.render(scale=scale).to_pil()
            image.save(out_dir / f"page-{index + 1}.png")
        print(f"Rendered {len(pdf)} page(s): {out_dir}")
    finally:
        pdf.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0)
    args = parser.parse_args()
    render(args.docx, args.out_dir, args.scale)
