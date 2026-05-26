from __future__ import annotations

import sys
from pathlib import Path

import fitz
from PIL import Image


DEFAULT_DPI = 200
JPEG_QUALITY = 95


def pause(message: str) -> None:
    if sys.stdin.isatty():
        input(message)


def select_pdfs() -> list[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return []

    root = tk.Tk()
    root.withdraw()
    root.update()
    paths = filedialog.askopenfilenames(
        title="Select PDF files",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()
    return [Path(path) for path in paths]


def render_pdf(pdf_path: Path, dpi: int = DEFAULT_DPI) -> Path:
    output_dir = pdf_path.with_name(f"{pdf_path.stem}_jpg")
    output_dir.mkdir(exist_ok=True)

    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    with fitz.open(pdf_path) as document:
        page_digits = max(3, len(str(document.page_count)))
        for index, page in enumerate(document, start=1):
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.frombytes(
                "RGB",
                (pixmap.width, pixmap.height),
                pixmap.samples,
            )
            image.save(
                output_dir / f"{pdf_path.stem}_{index:0{page_digits}d}.jpg",
                "JPEG",
                quality=JPEG_QUALITY,
                optimize=True,
            )

    return output_dir


def main() -> int:
    input_paths = [Path(arg).expanduser() for arg in sys.argv[1:]]
    if not input_paths:
        input_paths = select_pdfs()

    pdf_paths = [path for path in input_paths if path.is_file() and path.suffix.lower() == ".pdf"]
    if not pdf_paths:
        print("No PDF files selected.")
        pause("Press Enter to close...")
        return 1

    failures: list[tuple[Path, str]] = []
    for pdf_path in pdf_paths:
        try:
            output_dir = render_pdf(pdf_path.resolve())
            print(f"OK: {pdf_path.name} -> {output_dir}")
        except Exception as exc:
            failures.append((pdf_path, str(exc)))
            print(f"FAILED: {pdf_path} ({exc})")

    if failures:
        print()
        print("Some files could not be converted:")
        for pdf_path, message in failures:
            print(f"- {pdf_path}: {message}")

    print()
    pause("Done. Press Enter to close...")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
