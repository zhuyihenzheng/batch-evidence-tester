# -*- coding: utf-8 -*-
"""ファイル転送ができない環境へ、クリップボード経由でプロジェクトを持ち込むための
自己展開 PowerShell スクリプトを生成する。

想定シナリオ:
  AWS 上の Windows で、ブラウザもファイル転送も使えないが RDP のクリップボードは
  通る。テキストとして貼り付けられれば持ち込める、という状況。

方式:
  プロジェクトを ZIP → base64（ASCII のみ）→ 分割 → 各パートを PowerShell の
  ヒアストリングに埋め込む。base64 にすることで、文字コード・改行コード・
  日本語コメントの化けを完全に回避する。SHA-256 で完全性も検証する。

  python tools/make_transfer_bundle.py                # 既定 60KB 分割
  python tools/make_transfer_bundle.py --chunk-kb 30  # 貼り付け上限が厳しい場合
  python tools/make_transfer_bundle.py --minimal      # 実行に必要な最小構成のみ
"""

import argparse
import base64
import hashlib
import io
import os
import subprocess
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# --minimal のときに含めるもの。テスト・サンプル・fixtures を落とすと半分以下になる
MINIMAL_PREFIXES = ("src/", "config/", "tools/")
MINIMAL_FILES = ("check_env.py", "requirements.txt", "pyproject.toml",
                 "run_test.bat", "setup_windows.bat", "README.md")


def tracked_files():
    """git 管理下のファイル一覧。.gitignore の除外がそのまま効く。

    -z（NUL 区切り）は必須。既定の git ls-files は非 ASCII のパスを
    "cases/\\347\\222\\260..." のように 8 進エスケープ付きで引用して出すため、
    そのまま使うと日本語名のファイルが「存在しないパス」になって
    転送パックから黙って抜け落ちる。ケースをフォルダ名（＝タグ）で
    分ける運用では日本語名が普通に出てくるので、ここは必ず -z で受ける。
    """
    try:
        out = subprocess.check_output(["git", "-C", str(PROJECT_ROOT), "ls-files", "-z"])
    except (subprocess.CalledProcessError, OSError):
        print("git ls-files に失敗しました。git リポジトリ内で実行してください。", file=sys.stderr)
        raise SystemExit(1)
    return sorted(p for p in out.decode("utf-8").split("\0") if p)


def select_files(minimal):
    files = tracked_files()
    if not minimal:
        return files
    return [f for f in files
            if f.startswith(MINIMAL_PREFIXES) or f in MINIMAL_FILES]


def build_zip(files):
    """指定ファイルを ZIP にまとめてバイト列で返す。"""
    buf = io.BytesIO()
    # 圧縮率を上げて貼り付け量を減らす
    missing = []
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for rel in files:
            src = PROJECT_ROOT / rel
            if src.is_file():
                zf.write(src, arcname=rel)
            else:
                missing.append(rel)
    # 黙って落とすと、転送先で「なぜか動かない」まで気づけない
    if missing:
        print("[警告] 次のファイルを同梱できませんでした（%d 件）:" % len(missing), file=sys.stderr)
        for rel in missing:
            print("  - %s" % rel, file=sys.stderr)
    return buf.getvalue()


PART_TEMPLATE = """\
# ============================================================================
#  {project} 転送パック  {index} / {total}
#  貼り付けてそのまま実行してください（PowerShell）。
#  全パートを実行し終えたら unpack.ps1 を実行します。
# ============================================================================
$ErrorActionPreference = 'Stop'
$work = Join-Path $env:TEMP '{workdir}'
New-Item -ItemType Directory -Force -Path $work | Out-Null

$part = @'
{payload}
'@

Set-Content -Path (Join-Path $work '{partname}') -Value $part -Encoding Ascii -NoNewline
Write-Host ('パート {index}/{total} を保存しました: ' + (Join-Path $work '{partname}'))
"""

UNPACK_TEMPLATE = """\
# ============================================================================
#  {project} 展開スクリプト
#  全パートを実行した後に、これを貼り付けて実行してください。
#  展開先はカレントフォルダ配下の {project} です。
# ============================================================================
$ErrorActionPreference = 'Stop'
$work = Join-Path $env:TEMP '{workdir}'
$dest = Join-Path (Get-Location) '{project}'
$total = {total}

# --- パートの存在確認 -------------------------------------------------------
$missing = @()
for ($i = 1; $i -le $total; $i++) {{
    $name = 'part{{0:d2}}.txt' -f $i
    if (-not (Test-Path (Join-Path $work $name))) {{ $missing += $name }}
}}
if ($missing.Count -gt 0) {{
    Write-Host '次のパートが未実行です:' -ForegroundColor Red
    $missing | ForEach-Object {{ Write-Host ('  ' + $_) }}
    exit 1
}}

# --- 連結して base64 デコード ------------------------------------------------
$b64 = New-Object System.Text.StringBuilder
for ($i = 1; $i -le $total; $i++) {{
    $name = 'part{{0:d2}}.txt' -f $i
    [void]$b64.Append((Get-Content (Join-Path $work $name) -Raw))
}}
$text = $b64.ToString() -replace '\\s', ''
$bytes = [Convert]::FromBase64String($text)

$zipPath = Join-Path $work '{project}.zip'
[IO.File]::WriteAllBytes($zipPath, $bytes)

# --- 完全性の検証 -------------------------------------------------------------
$expected = '{sha256}'
$actual = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLower()
if ($actual -ne $expected) {{
    Write-Host 'SHA-256 が一致しません。貼り付けが途中で欠けた可能性があります。' -ForegroundColor Red
    Write-Host ('  期待値: ' + $expected)
    Write-Host ('  実際値: ' + $actual)
    exit 1
}}
Write-Host ('SHA-256 検証 OK: ' + $actual) -ForegroundColor Green

# --- 展開 ---------------------------------------------------------------------
if (Test-Path $dest) {{
    Write-Host ('既存フォルダを削除します: ' + $dest)
    Remove-Item -Recurse -Force $dest
}}
New-Item -ItemType Directory -Force -Path $dest | Out-Null

# Expand-Archive は PowerShell 5.0 以降。古い環境向けに .NET も用意する
if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {{
    Expand-Archive -Path $zipPath -DestinationPath $dest -Force
}} else {{
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $dest)
}}

$count = (Get-ChildItem -Recurse -File $dest).Count
Write-Host ''
Write-Host ('展開完了: ' + $dest + '  (' + $count + ' ファイル)') -ForegroundColor Green
Write-Host ''
Write-Host '次の手順:'
Write-Host ('  cd "' + $dest + '"')
Write-Host '  python check_env.py'

Remove-Item -Recurse -Force $work
"""


def main():
    parser = argparse.ArgumentParser(description="クリップボード転送用の自己展開スクリプトを生成する")
    parser.add_argument("--chunk-kb", type=int, default=60,
                        help="1 パートあたりの base64 サイズ（KB）。既定 60")
    parser.add_argument("--minimal", action="store_true",
                        help="テスト・サンプル・fixtures を除いた最小構成にする")
    parser.add_argument("--out", default="transfer", help="出力先フォルダ（既定: transfer）")
    parser.add_argument("--project", default="batch-evidence-tester", help="展開先フォルダ名")
    args = parser.parse_args()

    files = select_files(args.minimal)
    if not files:
        print("対象ファイルがありません。", file=sys.stderr)
        return 1

    raw = build_zip(files)
    sha256 = hashlib.sha256(raw).hexdigest()
    b64 = base64.b64encode(raw).decode("ascii")

    # 76 文字ごとに改行を入れる（ヒアストリング内で扱いやすく、目視でも追える）
    wrapped = "\n".join(b64[i:i + 76] for i in range(0, len(b64), 76))

    chunk_size = args.chunk_kb * 1024
    # 行単位で切る（行の途中で切ると連結ミスに気づきにくい）
    lines = wrapped.split("\n")
    parts, current, current_len = [], [], 0
    for line in lines:
        if current_len + len(line) + 1 > chunk_size and current:
            parts.append("\n".join(current))
            current, current_len = [], 0
        current.append(line)
        current_len += len(line) + 1
    if current:
        parts.append("\n".join(current))

    out_dir = PROJECT_ROOT / args.out
    out_dir.mkdir(parents=True, exist_ok=True)
    workdir = "bet_transfer"

    for i, payload in enumerate(parts, start=1):
        script = PART_TEMPLATE.format(
            project=args.project, index=i, total=len(parts),
            workdir=workdir, partname="part%02d.txt" % i, payload=payload,
        )
        (out_dir / ("paste_%02d.ps1" % i)).write_text(script, encoding="utf-8")

    (out_dir / "unpack.ps1").write_text(
        UNPACK_TEMPLATE.format(project=args.project, workdir=workdir,
                               total=len(parts), sha256=sha256),
        encoding="utf-8")

    print("生成しました: %s" % out_dir)
    print("  対象ファイル : %d 件%s" % (len(files), "（最小構成）" if args.minimal else ""))
    print("  ZIP          : %.1f KB" % (len(raw) / 1024))
    print("  base64       : %.1f KB" % (len(b64) / 1024))
    print("  SHA-256      : %s" % sha256)
    print("  パート数     : %d（1 パート最大 %d KB）" % (len(parts), args.chunk_kb))
    print()
    print("Windows 側の手順:")
    for i in range(1, len(parts) + 1):
        print("  %d. paste_%02d.ps1 の中身を PowerShell に貼り付けて実行" % (i, i))
    print("  %d. unpack.ps1 の中身を貼り付けて実行（SHA-256 検証 + 展開）" % (len(parts) + 1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
