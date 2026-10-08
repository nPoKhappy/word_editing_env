"""Export DOCX with installed Microsoft Word and render its PDF pages to PNG."""

import argparse
from pathlib import Path

from word_session import WordSession, run_bounded


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
        for index in range(len(pdf)):
            page = pdf[index]
            image = page.render(scale=scale).to_pil()
            image.save(out_dir / f"page-{index + 1}.png")
            page.close()
        print(f"Rendered {len(pdf)} page(s): {out_dir}")
    finally:
        pdf.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0)
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--worker-state", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker_state:
        render(args.docx, args.out_dir, args.scale, args.worker_state)
    else:
        import sys
        sys.exit(run_bounded(Path(__file__), sys.argv[1:], args.timeout))
