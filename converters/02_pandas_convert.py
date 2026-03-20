"""
変換ツール02: pandas
pd.read_excel() で各シートをDataFrameとして読み込み、
to_markdown() でMarkdownテーブルに変換する。
方眼紙Excelとの相性：結合セルはNaNで埋まるため素直には変換できない。
"""

import time
from pathlib import Path

import pandas as pd

INPUT_FILE = Path("test_data/houganshi_sample.xlsx")
OUTPUT_DIR = Path("output/02_pandas")


def convert(input_file: Path, output_dir: Path):
    start = time.perf_counter()

    xl = pd.ExcelFile(input_file)
    result_parts = [f"# pandas 変換結果\n\n**入力ファイル:** `{input_file}`\n"]

    for sheet_name in xl.sheet_names:
        result_parts.append(f"## シート: {sheet_name}\n")
        try:
            df = pd.read_excel(
                input_file,
                sheet_name=sheet_name,
                header=None,   # ヘッダー自動判定なし（方眼紙は1行目がヘッダーとは限らない）
            )
            # 全NaN行・列を除去
            df = df.dropna(how="all").dropna(axis=1, how="all")
            df = df.fillna("")

            if df.empty:
                result_parts.append("*（データなし）*\n")
                continue

            # 値を文字列化
            df = df.astype(str).replace("nan", "").replace("<NA>", "")

            md_table = df.to_markdown(index=False, headers="keys")
            result_parts.append(md_table)
            result_parts.append("")

        except Exception as e:
            result_parts.append(f"*エラー: {e}*\n")

    elapsed = time.perf_counter() - start
    result_parts.append(f"\n---\n*変換時間: {elapsed:.3f}秒 | ツール: pandas {pd.__version__}*\n")

    md_text = "\n".join(result_parts)
    out_path = output_dir / "result.md"
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path.write_text(md_text, encoding="utf-8")

    print(f"[02_pandas] 完了 ({elapsed:.3f}秒) -> {out_path}")
    return elapsed, md_text


if __name__ == "__main__":
    convert(INPUT_FILE, OUTPUT_DIR)
