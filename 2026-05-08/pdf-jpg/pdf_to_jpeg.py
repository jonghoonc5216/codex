from __future__ import annotations

import sys
from pathlib import Path


DEFAULT_DPI = 200
JPEG_QUALITY = 95

try:
    import fitz
    from PIL import Image
except ModuleNotFoundError as exc:
    fitz = None
    Image = None
    MISSING_MODULE = exc.name
else:
    MISSING_MODULE = ""


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
        title="JPEG로 변환할 PDF 파일 선택",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()
    return [Path(path) for path in paths]


def expand_inputs(input_paths: list[Path]) -> list[Path]:
    pdf_paths: list[Path] = []
    for path in input_paths:
        if path.is_file() and path.suffix.lower() == ".pdf":
            pdf_paths.append(path)
        elif path.is_dir():
            pdf_paths.extend(sorted(path.glob("*.pdf")))
    return pdf_paths


def render_pdf(pdf_path: Path, dpi: int = DEFAULT_DPI) -> Path:
    if fitz is None or Image is None:
        raise RuntimeError("필요 라이브러리가 설치되어 있지 않습니다.")

    output_dir = pdf_path.with_name(f"{pdf_path.stem}_jpeg")
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
            jpg_path = output_dir / f"{pdf_path.stem}_{index:0{page_digits}d}.jpg"
            image.save(jpg_path, "JPEG", quality=JPEG_QUALITY, optimize=True)
            print(f"  - {jpg_path.name}")

    return output_dir


def main() -> int:
    if MISSING_MODULE:
        print("PDF 변환에 필요한 Python 라이브러리가 없습니다.")
        print(f"누락된 항목: {MISSING_MODULE}")
        print("같은 폴더의 '1_처음한번_필요라이브러리_설치.bat'을 먼저 실행해 주세요.")
        pause("Enter 키를 누르면 닫힙니다...")
        return 2

    input_paths = [Path(arg).expanduser() for arg in sys.argv[1:]]
    if not input_paths:
        input_paths = select_pdfs()

    pdf_paths = expand_inputs(input_paths)
    if not pdf_paths:
        print("선택된 PDF 파일이 없습니다.")
        pause("Enter 키를 누르면 닫힙니다...")
        return 1

    failures: list[tuple[Path, str]] = []
    for pdf_path in pdf_paths:
        try:
            print(f"변환 중: {pdf_path}")
            output_dir = render_pdf(pdf_path.resolve())
            print(f"완료: {output_dir}")
        except Exception as exc:
            failures.append((pdf_path, str(exc)))
            print(f"실패: {pdf_path} ({exc})")

    if failures:
        print()
        print("변환하지 못한 파일:")
        for pdf_path, message in failures:
            print(f"- {pdf_path}: {message}")

    print()
    pause("작업이 끝났습니다. Enter 키를 누르면 닫힙니다...")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
