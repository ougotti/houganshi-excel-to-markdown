"""
変換ツール04b: docling 改善版

【デフォルト版の問題】
- gap_tolerance=0 のため、方眼紙の「空白列/行で区切られた各ブロック」が
  細かく分割されすぎる or 逆にひとつの巨大テーブルになる
- treat_singleton_as_text=False のため、単独セル（タイトルや注記）も
  Tableとして扱われノイズになる
- Markdown出力では row_span/col_span が表現できないため情報落ちがある

【改善内容】
1. MsExcelBackendOptions で方眼紙向けにチューニング
   - gap_tolerance=1  : 空白1行/列まで同一テーブルとして扱う
   - treat_singleton_as_text=True : 単独セルをテキスト扱いにしてノイズ削減
2. Markdown に加えて HTML も出力（rowspan/colspan が保持される）
3. TableItem を個別に取り出し、各テーブルのDataFrameも保存
"""

import json
import time
from pathlib import Path

from docling.datamodel.backend_options import MsExcelBackendOptions
from docling.datamodel.base_models import InputFormat
from docling.document_converter import DocumentConverter, ExcelFormatOption
from docling_core.types.doc import TableItem, TextItem

INPUT_FILE = Path("test_data/houganshi_sample.xlsx")
OUTPUT_DIR = Path("output/04b_docling_improved")


def convert(input_file: Path, output_dir: Path):
    start = time.perf_counter()

    # 方眼紙向けオプション
    backend_options = MsExcelBackendOptions(
        gap_tolerance=1,               # 空白行/列1つまで同一テーブルと見なす
        treat_singleton_as_text=True,  # 孤立セルをテキスト扱い（タイトル行など）
    )

    converter = DocumentConverter(
        allowed_formats=[InputFormat.XLSX],
        format_options={
            InputFormat.XLSX: ExcelFormatOption(
                backend_options=backend_options
            )
        },
    )

    doc_result = converter.convert(str(input_file))
    doc = doc_result.document

    elapsed = time.perf_counter() - start

    import docling
    version = getattr(docling, "__version__", "unknown")

    output_dir.mkdir(parents=True, exist_ok=True)

    # --- 1. Markdown 出力 ---
    md_text_raw = doc.export_to_markdown()
    footer = (
        f"\n---\n"
        f"*変換時間: {elapsed:.3f}秒 | "
        f"ツール: docling {version} "
        f"(gap_tolerance=1, treat_singleton_as_text=True)*\n"
    )
    full_md = (
        f"# docling 改善版 変換結果\n\n"
        f"**入力ファイル:** `{input_file}`  \n"
        f"**改善内容:** `gap_tolerance=1`, `treat_singleton_as_text=True`\n\n"
        + md_text_raw
        + footer
    )
    md_path = output_dir / "result.md"
    md_path.write_text(full_md, encoding="utf-8")

    # --- 2. HTML 出力（rowspan/colspan が保持される） ---
    try:
        html_text = doc.export_to_html()
        html_path = output_dir / "result.html"
        html_path.write_text(html_text, encoding="utf-8")
        html_note = f"HTML出力: {html_path}"
    except Exception as e:
        html_note = f"HTML出力エラー: {e}"

    # --- 3. テーブル一覧を個別に出力 ---
    table_lines = [
        "# docling 改善版: テーブル別出力\n",
        f"**入力ファイル:** `{input_file}`\n",
        f"**オプション:** gap_tolerance=1, treat_singleton_as_text=True\n",
    ]
    table_count = 0
    text_count = 0

    for item, level in doc.iterate_items():
        if isinstance(item, TableItem):
            table_count += 1
            table_lines.append(f"\n## テーブル {table_count}\n")
            # Markdown テーブル
            try:
                table_lines.append(item.export_to_markdown())
            except Exception as e:
                table_lines.append(f"*Markdownエラー: {e}*")

            # DataFrame 情報（行列数）
            try:
                df = item.export_to_dataframe()
                table_lines.append(
                    f"\n> DataFrame: {df.shape[0]}行 × {df.shape[1]}列"
                )
            except Exception:
                pass

            # HTML（rowspan/colspan あり）
            try:
                table_lines.append("\n**HTML（rowspan/colspan保持）:**")
                table_lines.append(f"```html\n{item.export_to_html()}\n```")
            except Exception as e:
                table_lines.append(f"*HTMLエラー: {e}*")

        elif isinstance(item, TextItem):
            text_count += 1
            table_lines.append(f"\n**テキスト:** {item.text}")

    tables_path = output_dir / "tables.md"
    tables_path.write_text("\n".join(table_lines), encoding="utf-8")

    print(f"[04b_docling_improved] 完了 ({elapsed:.3f}秒)")
    print(f"  Markdown -> {md_path}")
    print(f"  {html_note}")
    print(f"  テーブル別 -> {tables_path}")
    print(f"  検出: テーブル {table_count}個, テキスト {text_count}個")
    return elapsed, full_md


if __name__ == "__main__":
    convert(INPUT_FILE, OUTPUT_DIR)
