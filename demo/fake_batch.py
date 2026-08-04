"""検証用のダミー batch。実際の C# 製 .exe の代わりに動く。

本物の batch と同じ「入出力の作法」だけを再現する:
  投入フォルダの CSV を読む → 出力フォルダに結果 CSV → 処理済/エラーへ移動 → ログ追記
これにより、実環境に触れる前にツール側（証跡採取・判定・Excel 出力）を検証できる。

  python demo/fake_batch.py --mode daily --date 20260803 --root ./sandbox
"""

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import List


def log(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} [INFO] {message}\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", default="daily")
    parser.add_argument("--date", default=datetime.now().strftime("%Y%m%d"))
    parser.add_argument("--root", default="./sandbox")
    args = parser.parse_args()

    root = Path(args.root)
    in_dir, processed_dir = root / "in", root / "processed"
    error_dir, out_dir, log_dir = root / "error", root / "out", root / "log"
    for d in (in_dir, processed_dir, error_dir, out_dir, log_dir):
        d.mkdir(parents=True, exist_ok=True)

    log_path = log_dir / f"batch_{args.date}.log"
    log(log_path, f"処理開始 mode={args.mode} date={args.date}")

    targets = sorted(in_dir.glob("*.csv"))
    log(log_path, f"対象ファイル数={len(targets)}")

    total_ok = total_ng = 0
    results: List[List[str]] = []

    for src in targets:
        log(log_path, f"ファイル取込開始 {src.name}")
        rows_ok = rows_ng = 0
        with src.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for line_no, row in enumerate(reader, start=2):
                order_no = (row.get("ORDER_NO") or "").strip()
                amount_raw = (row.get("AMOUNT") or "").strip()
                try:
                    amount = int(amount_raw)
                    if amount <= 0:
                        raise ValueError("金額が 0 以下です")
                except ValueError as exc:
                    rows_ng += 1
                    print(f"[ERROR] {src.name}:{line_no} {order_no} {exc}", file=sys.stderr)
                    with log_path.open("a", encoding="utf-8") as lf:
                        lf.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} [ERROR] 明細エラー {order_no} {exc}\n")
                    continue
                rows_ok += 1
                results.append([order_no, "完了", f"{amount}"])

        total_ok += rows_ok
        total_ng += rows_ng
        dest = (error_dir if rows_ng else processed_dir) / src.name
        src.replace(dest)
        log(log_path, f"ファイル取込終了 {src.name} 正常={rows_ok} エラー={rows_ng} 移動先={dest.parent.name}")

    if results:
        out_path = out_dir / f"RESULT_{args.date}.csv"
        with out_path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["ORDER_NO", "STATUS", "AMOUNT"])
            writer.writerows(results)
        log(log_path, f"結果ファイル出力 {out_path.name} 件数={len(results)}")

    exit_code = 1 if total_ng else 0
    log(log_path, f"取込件数={total_ok}")
    log(log_path, "処理正常終了" if exit_code == 0 else "処理異常終了")
    print(f"取込件数={total_ok} エラー件数={total_ng}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
