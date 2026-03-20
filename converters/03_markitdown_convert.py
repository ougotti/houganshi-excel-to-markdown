"""
変換ツール03: markitdown (Microsoft)
MarkItDown は Microsoft 製の汎用ドキュメント→Markdown 変換ライブラリ。
Excel (.xlsx) にも対応している。
"""

import time
from pathlib import Path

from markitdown import MarkItDown

INPUT_FILE = Path("test_data/houganshi_sample.xlsx")
OUTPUT_DIR = Path("output/03_markitdown")


def convert(input_file: Path, output_dir: Path):
    start = time.perf_counter()

    md_converter = MarkItDown()
    result = md_converter.convert(str(input_file))
    md_text_raw = result.text_content

    elapsed = time.perf_counter() - start

    import markitdown
    footer = f"\n---\n*変換時間: {elapsed:.3f}秒 | ツール: markitdown {markitdown.__version__}*\n"
    full_text = (
        f"# markitdown 変換結果\n\n**入力ファイル:** `{input_file}`\n\n"
        + md_text_raw
        + footer
    )

    out_path = output_dir / "result.md"
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path.write_text(full_text, encoding="utf-8")

    print(f"[03_markitdown] 完了 ({elapsed:.3f}秒) -> {out_path}")
    return elapsed, full_text


if __name__ == "__main__":
    convert(INPUT_FILE, OUTPUT_DIR)
