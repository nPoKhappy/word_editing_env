"""Export DOCX with installed Microsoft Word and render its PDF pages to PNG."""

import argparse
from pathlib import Path
import subprocess
import sys

from word_session import WordSession, run_bounded


def render_pdf_page(pdf_path: Path, out_dir: Path, scale: float, index: int) -> None:
    """Own the bitmap through save; isolate PDFium font/cache state per page."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf[index]
        bitmap = page.render(scale=scale)
        try:
            image = bitmap.to_pil()
            try:
                image.save(out_dir / f"page-{index + 1}.png")
            finally:
                image.close()
        finally:
            bitmap.close()
            page.close()
    finally:
        pdf.close()


def render(docx_path: Path, out_dir: Path, scale: float, state_path=None) -> None:
    if not docx_path.is_file():
        raise FileNotFoundError(docx_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    import pypdfium2 as pdfium

    pdf_path = out_dir / (docx_path.stem + ".pdf")
    with WordSession(docx_path, read_only=True, state_path=state_path) as session:
        session.document.ExportAsFixedFormat(str(pdf_path.resolve()), 17)

    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page_count = len(pdf)
    finally:
        pdf.close()
    # A 61-page Word export exposed missing glyphs in a late PNG despite a
    # correct PDF. Fresh single-page rendering reproduced all text correctly.
    # Keep page workers bounded and avoid carrying native renderer state forward.
    for index in range(page_count):
        subprocess.run([sys.executable, "-X", "utf8", str(Path(__file__).resolve()),
                        str(pdf_path), "--out-dir", str(out_dir), "--scale", str(scale),
                        "--pdf-page", str(index)], check=True, timeout=30,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    print(f"Rendered {page_count} page(s), isolated PDFium pages: {out_dir}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0)
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--worker-state", help=argparse.SUPPRESS)
    parser.add_argument("--pdf-page", type=int, help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.pdf_page is not None:
        render_pdf_page(args.docx, args.out_dir, args.scale, args.pdf_page)
    elif args.worker_state:
        render(args.docx, args.out_dir, args.scale, args.worker_state)
    else:
        sys.exit(run_bounded(Path(__file__), sys.argv[1:], args.timeout))
