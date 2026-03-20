"""
変換ツール04: docling (IBM Research)
Docling は IBM 製の高精度ドキュメントパーサー。
レイアウト解析・テーブル検出が強みで、Excel も対応。
"""

import time
from pathlib import Path

INPUT_FILE = Path("test_data/houganshi_sample.xlsx")
OUTPUT_DIR = Path("output/04_docling")


def convert(input_file: Path, output_dir: Path):
    start = time.perf_counter()

    try:
        from docling.document_converter import DocumentConverter
        converter = DocumentConverter()
        doc_result = converter.convert(str(input_file))
        md_text_raw = doc_result.document.export_to_markdown()

        import docling
        version = getattr(docling, "__version__", "unknown")
    except Exception as e:
        md_text_raw = f"*変換エラー: {e}*\n"
        version = "unknown"

    elapsed = time.perf_counter() - start

    footer = f"\n---\n*変換時間: {elapsed:.3f}秒 | ツール: docling {version}*\n"
    full_text = (
        f"# docling 変換結果\n\n**入力ファイル:** `{input_file}`\n\n"
        + md_text_raw
        + footer
    )

    out_path = output_dir / "result.md"
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path.write_text(full_text, encoding="utf-8")

    print(f"[04_docling] 完了 ({elapsed:.3f}秒) -> {out_path}")
    return elapsed, full_text


if __name__ == "__main__":
    convert(INPUT_FILE, OUTPUT_DIR)
