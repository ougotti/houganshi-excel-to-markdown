"""
全変換ツールを一括実行し、結果をまとめる比較ランナー。

実行方法:
    .venv/Scripts/python run_all.py

出力:
    output/XX_toolname/result.md  ... 各ツールの変換結果
    output/comparison.md          ... ツール比較サマリー
"""

import importlib
import sys
import time
from pathlib import Path

INPUT_FILE = Path("test_data/houganshi_sample.xlsx")
OUTPUT_DIR  = Path("output")

CONVERTERS = [
    ("01_openpyxl",   "converters.01_openpyxl_convert"),
    ("02_pandas",     "converters.02_pandas_convert"),
    ("03_markitdown", "converters.03_markitdown_convert"),
    ("04_docling",    "converters.04_docling_convert"),
]


def count_stats(md_text: str) -> dict:
    """Markdownテキストの簡易統計を返す"""
    lines = md_text.splitlines()
    table_rows  = sum(1 for l in lines if l.strip().startswith("|"))
    headings    = sum(1 for l in lines if l.strip().startswith("#"))
    total_chars = len(md_text)
    total_lines = len(lines)
    return {
        "文字数":     total_chars,
        "行数":       total_lines,
        "テーブル行": table_rows,
        "見出し数":   headings,
    }


def main():
    if not INPUT_FILE.exists():
        print(f"エラー: 入力ファイルが見つかりません: {INPUT_FILE}")
        print("先に create_test_data.py を実行してください。")
        sys.exit(1)

    print(f"入力ファイル: {INPUT_FILE}")
    print("=" * 60)

    results = {}  # tool_name -> {elapsed, stats, error}

    for tool_name, module_path in CONVERTERS:
        print(f"\n>> {tool_name} を実行中...")
        out_dir = OUTPUT_DIR / tool_name
        try:
            mod = importlib.import_module(module_path)
            elapsed, md_text = mod.convert(INPUT_FILE, out_dir)
            stats = count_stats(md_text)
            results[tool_name] = {"elapsed": elapsed, "stats": stats, "error": None}
        except ModuleNotFoundError as e:
            msg = f"パッケージ未インストール: {e}"
            print(f"  [SKIP] {msg}")
            results[tool_name] = {"elapsed": None, "stats": None, "error": msg}
        except Exception as e:
            msg = str(e)
            print(f"  [ERROR] {msg}")
            results[tool_name] = {"elapsed": None, "stats": None, "error": msg}

    # --- 比較サマリーを生成 ---
    print("\n" + "=" * 60)
    print("比較サマリーを生成中...")

    lines = [
        "# 変換ツール比較サマリー",
        "",
        f"**入力ファイル:** `{INPUT_FILE}`  ",
        "",
        "## 結果一覧",
        "",
        "| ツール | 変換時間(秒) | 文字数 | 行数 | テーブル行 | 見出し数 | 備考 |",
        "| --- | ---: | ---: | ---: | ---: | ---: | --- |",
    ]

    for tool_name, res in results.items():
        if res["error"]:
            lines.append(f"| {tool_name} | - | - | - | - | - | ⚠️ {res['error']} |")
        else:
            s = res["stats"]
            t = f"{res['elapsed']:.3f}"
            lines.append(
                f"| {tool_name} | {t} | {s['文字数']:,} | {s['行数']:,} | "
                f"{s['テーブル行']:,} | {s['見出し数']} | ✅ |"
            )

    lines += [
        "",
        "## 各ツールの特徴メモ",
        "",
        "| ツール | 方眼紙対応 | 画像 | グラフ | 処理速度 | 導入難度 |",
        "| --- | :---: | :---: | :---: | :---: | :---: |",
        "| openpyxl | △ 結合セル要工夫 | ✗ | ✗ | ⚡ 速い | 低 |",
        "| pandas | ✗ NaNが残る | ✗ | ✗ | ⚡ 速い | 低 |",
        "| markitdown | △ | ✗ | ✗ | ⚡ 速い | 低 |",
        "| docling | △ | △ | ✗ | 🐢 遅い | 高 |",
        "",
        "---",
        f"*生成日時: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*",
        "",
    ]

    summary_path = OUTPUT_DIR / "comparison.md"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    summary_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"比較サマリー -> {summary_path}")
    print("=" * 60)
    print("完了！")


if __name__ == "__main__":
    main()
