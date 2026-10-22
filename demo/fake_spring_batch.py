"""検証用のダミー Spring Boot batch。実際の実行可能 jar の代わりに動く。

本物の Spring Boot + Spring Batch と同じ「外から見える作法」だけを再現する:
  - 起動: java [JVM オプション] -jar app.jar --spring.batch.job.name=<ジョブ> key=value ...
    （settings の type: java が組み立てるコマンドをそのまま受け取る。-jar / -D... は読み捨てる）
  - ログ: Spring Boot 3 の既定書式。ジョブランチャの launched / completed 行を出す
  - 終了コード: COMPLETED=0 / FAILED=1（main で System.exit(SpringApplication.exit(...)) する作り）
処理内容は demo/fake_batch.py と同じ（投入 CSV を読み、結果 CSV・処理済/エラー移動）。

検証用の追加オプション（本物には無い）:
  --fake.root=<dir>                 入出力フォルダの親（in / out / processed / error）
  --fake.exit-code=always-zero      失敗しても 0 で終わる作り（よくある不備）を再現する
  --fake.job-repository=<json>      完了済みジョブの記録。同じパラメータの再実行を
                                    JobInstanceAlreadyCompleteException で拒否する

  python demo/fake_spring_batch.py -jar app.jar --spring.batch.job.name=orderImportJob targetDate=20260803
"""

import csv
import io
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

PID = os.getpid()
LAUNCHER = "o.s.b.c.l.s.TaskExecutorJobLauncher"


class Logger:
    def __init__(self, log_path):
        self.log_path = log_path

    def __call__(self, level, logger, message):
        stamp = datetime.now().astimezone().isoformat(timespec="milliseconds")
        line = "%s %5s %d --- [order-batch] [           main] %-40s : %s" % (
            stamp, level, PID, logger, message)
        print(line, flush=True)
        if self.log_path:
            self.log_path.parent.mkdir(parents=True, exist_ok=True)
            with self.log_path.open("a", encoding="utf-8") as f:
                f.write(line + "\n")


def parse_args(argv):
    options, params = {}, {}
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "-jar":
            i += 2           # 実行する jar 自身。ここでは読み捨てる
            continue
        if arg.startswith("-D") or arg.startswith("-X"):
            i += 1           # JVM オプション
            continue
        if arg.startswith("--") and "=" in arg:
            key, value = arg[2:].split("=", 1)
            options[key] = value
        elif "=" in arg and not arg.startswith("-"):
            key, value = arg.split("=", 1)
            params[key] = value   # ジョブパラメータ（identifying）
        i += 1
    return options, params


def describe(params):
    return "[{%s}]" % ", ".join(
        "'%s':'{value=%s, type=class java.lang.String, identifying=true}'" % (k, v)
        for k, v in sorted(params.items()))


def main():
    # 本物では -Dstdout.encoding=UTF-8 を付ける想定。console_encoding: utf-8 と揃える
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    options, params = parse_args(sys.argv[1:])
    job = options.get("spring.batch.job.name") or options.get("spring.batch.job.names") or "orderImportJob"
    log_file = options.get("logging.file.name")
    log = Logger(Path(log_file) if log_file else None)
    root = Path(options.get("fake.root", "./sandbox_spring"))
    always_zero = options.get("fake.exit-code") == "always-zero"
    started = time.time()

    log("INFO", "c.e.batch.OrderBatchApplication", "Starting OrderBatchApplication using Java 17 with PID %d" % PID)
    log("INFO", "c.e.batch.OrderBatchApplication", "Started OrderBatchApplication in 1.234 seconds")

    repository = options.get("fake.job-repository")
    record = "%s %s" % (job, json.dumps(params, sort_keys=True))
    if repository and Path(repository).is_file():
        done = json.loads(Path(repository).read_text(encoding="utf-8"))
        if record in done:
            log("ERROR", "o.s.boot.SpringApplication", "Application run failed")
            print("org.springframework.batch.core.repository.JobInstanceAlreadyCompleteException: "
                  "A job instance already exists and is complete for identifying parameters=%s.  "
                  "If you want to run this job again, change the parameters." % describe(params), flush=True)
            return 0 if always_zero else 1

    log("INFO", LAUNCHER, "Job: [SimpleJob: [name=%s]] launched with the following parameters: %s"
        % (job, describe(params)))
    in_dir, processed_dir = root / "in", root / "processed"
    error_dir, out_dir = root / "error", root / "out"

    status, ok_rows, results = "COMPLETED", 0, []
    if not in_dir.is_dir() or not out_dir.is_dir():
        log("ERROR", "o.s.batch.core.step.AbstractStep", "Encountered an error executing step importStep in job %s" % job)
        print("java.lang.IllegalStateException: 入力または出力フォルダが存在しません: %s" % root, flush=True)
        status = "FAILED"
    else:
        for src in sorted(in_dir.glob("*.csv")):
            ng = 0
            with src.open("r", encoding="utf-8-sig", newline="") as f:
                for row in csv.DictReader(f):
                    try:
                        amount = int((row.get("AMOUNT") or "").strip())
                        if amount <= 0:
                            raise ValueError("金額が 0 以下です")
                    except ValueError as exc:
                        ng += 1
                        log("ERROR", "c.e.batch.OrderItemProcessor", "明細エラー %s %s" % (row.get("ORDER_NO"), exc))
                        continue
                    ok_rows += 1
                    results.append([row.get("ORDER_NO"), "完了", str(amount)])
            target = error_dir if ng else processed_dir
            target.mkdir(parents=True, exist_ok=True)
            src.replace(target / src.name)
            if ng:
                status = "FAILED"
        if status == "COMPLETED" and results:
            out_path = out_dir / ("RESULT_%s.csv" % params.get("targetDate", datetime.now().strftime("%Y%m%d")))
            with out_path.open("w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["ORDER_NO", "STATUS", "AMOUNT"])
                writer.writerows(results)
        log("INFO", "o.s.batch.core.step.AbstractStep", "Step: [importStep] executed in %dms" % ((time.time() - started) * 1000))
        log("INFO", "c.e.batch.OrderItemWriter", "取込件数=%d" % (ok_rows if status == "COMPLETED" else 0))

    log("INFO", LAUNCHER, "Job: [SimpleJob: [name=%s]] completed with the following parameters: %s "
        "and the following status: [%s] in %dms" % (job, describe(params), status, (time.time() - started) * 1000))
    if repository and status == "COMPLETED":
        path = Path(repository)
        done = json.loads(path.read_text(encoding="utf-8")) if path.is_file() else []
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(done + [record], ensure_ascii=False), encoding="utf-8")
    if always_zero:
        return 0
    return 0 if status == "COMPLETED" else 1


if __name__ == "__main__":
    sys.exit(main())
