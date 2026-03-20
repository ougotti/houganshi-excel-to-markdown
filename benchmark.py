"""
各変換ツールの処理時間を N 回計測して平均・最小・最大を出すベンチマーク。

注意:
  - docling は初回実行時にモデルをキャッシュする。
    1回目（コールドスタート）と2回目以降（ウォームスタート）を区別して記録する。
  - 他ツールも同様にコールド/ウォームの差を確認できる。

実行方法:
    .venv/Scripts/python benchmark.py [--runs N]
"""

import argparse
import importlib
import statistics
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


def run_benchmark(runs: int = 5):
    if not INPUT_FILE.exists():
        print(f"エラー: {INPUT_FILE} が見つかりません。先に create_test_data.py を実行してください。")
        sys.exit(1)

    print(f"ベンチマーク開始: {runs} 回計測")
    print(f"入力ファイル: {INPUT_FILE}")
    print("=" * 70)

    all_results = {}

    for tool_name, module_path in CONVERTERS:
        print(f"\n>> {tool_name}")
        out_dir = OUTPUT_DIR / tool_name
        times = []

        try:
            mod = importlib.import_module(module_path)
        except ModuleNotFoundError as e:
            print(f"   [SKIP] {e}")
            all_results[tool_name] = None
            continue

        for i in range(runs):
            label = "コールド" if i == 0 else f"ウォーム{i}"
            t0 = time.perf_counter()
            try:
                mod.convert(INPUT_FILE, out_dir)
                elapsed = time.perf_counter() - t0
                times.append(elapsed)
                print(f"   [{i+1}/{runs}] {label}: {elapsed:.3f} 秒")
            except Exception as e:
                print(f"   [{i+1}/{runs}] エラー: {e}")

        if times:
            cold  = times[0]
            warm  = times[1:] if len(times) > 1 else times
            all_results[tool_name] = {
                "cold":  cold,
                "warm_avg": statistics.mean(warm),
                "warm_min": min(warm),
                "warm_max": max(warm),
                "all":   times,
            }
            print(f"   --- コールド: {cold:.3f}s | ウォーム平均: {statistics.mean(warm):.3f}s "
                  f"(min {min(warm):.3f}s / max {max(warm):.3f}s)")

    # --- サマリー出力 ---
    print("\n" + "=" * 70)
    print("サマリー")
    print("=" * 70)

    header = f"{'ツール':<20} {'コールド(秒)':>12} {'ウォーム平均(秒)':>16} {'min':>8} {'max':>8}"
    print(header)
    print("-" * 70)

    summary_rows = []
    for tool_name, res in all_results.items():
        if res is None:
            print(f"{tool_name:<20} {'SKIP':>12}")
            summary_rows.append((tool_name, None, None, None, None))
        else:
            print(
                f"{tool_name:<20} {res['cold']:>12.3f} {res['warm_avg']:>16.3f} "
                f"{res['warm_min']:>8.3f} {res['warm_max']:>8.3f}"
            )
            summary_rows.append((
                tool_name,
                res["cold"],
                res["warm_avg"],
                res["warm_min"],
                res["warm_max"],
            ))

    # --- Markdown 保存 ---
    md_lines = [
        "# ベンチマーク結果",
        "",
        f"**入力ファイル:** `{INPUT_FILE}`  ",
        f"**計測回数:** {runs} 回  ",
        "",
        "## 処理時間（秒）",
        "",
        "| ツール | コールド | ウォーム平均 | ウォーム min | ウォーム max |",
        "| --- | ---: | ---: | ---: | ---: |",
    ]
    for tool_name, cold, warm_avg, warm_min, warm_max in summary_rows:
        if cold is None:
            md_lines.append(f"| {tool_name} | SKIP | - | - | - |")
        else:
            md_lines.append(
                f"| {tool_name} | {cold:.3f} | {warm_avg:.3f} | {warm_min:.3f} | {warm_max:.3f} |"
            )

    md_lines += [
        "",
        "## 備考",
        "",
        "- **コールド**: 初回実行（モジュール初期化・モデルキャッシュ読込を含む）",
        "- **ウォーム**: 2回目以降（モデル・モジュールがメモリ/キャッシュに乗った状態）",
        "- docling は初回に HuggingFace モデルをロードするためコールドが特に遅い",
        "",
        f"*計測日時: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*",
        "",
    ]

    out_path = OUTPUT_DIR / "benchmark.md"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"\nベンチマーク結果 -> {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--runs", type=int, default=5, help="計測回数 (デフォルト: 5)")
    args = parser.parse_args()
    run_benchmark(args.runs)
