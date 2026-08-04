# ============================================================================
#  隔離環境の Windows へ Git / VS Code をブラウザ無しで導入する
#
#  想定: AWS 上の Windows Server。ブラウザもファイル転送も使えないが、
#        HTTPS の外向き通信だけは通る、という状況。
#
#  使い方（PowerShell に貼り付けて実行）:
#      .\install_tools_windows.ps1 -What git          # Git だけ
#      .\install_tools_windows.ps1 -What code         # VS Code だけ
#      .\install_tools_windows.ps1 -What all          # 両方
#      .\install_tools_windows.ps1 -What all -NoAdmin # 管理者権限なしで導入
#
#  -NoAdmin を付けると:
#      Git   -> PortableGit（インストール不要の自己展開版）
#      Code  -> User Installer（管理者権限不要。%LOCALAPPDATA% へ入る）
#  どちらも管理者権限もレジストリ変更も要らないので、権限が下りない環境向け。
# ============================================================================
[CmdletBinding()]
param(
    [ValidateSet('git', 'code', 'all')]
    [string]$What = 'all',
    [switch]$NoAdmin,
    [string]$InstallRoot = 'C:\tools'
)

$ErrorActionPreference = 'Stop'

# --- 古い Windows での 2 大ハマりどころを先に潰す -----------------------------
# 1) Windows Server 2016 等は既定で TLS 1.2 が無効。有効化しないと HTTPS が全滅する
[Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
# 2) Invoke-WebRequest は既定で IE のエンジンを使うため、IE 初回起動設定が
#    未完了だと例外になる。-UseBasicParsing を必ず付ける（下の Get-File 内で対応）

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-File {
    param([string]$Uri, [string]$OutFile)
    Write-Host ("  取得中: " + $Uri)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    # 進捗表示を切ると Invoke-WebRequest が体感で数倍速くなる（既知の挙動）
    $prev = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing
    } finally {
        $ProgressPreference = $prev
    }
    $mb = (Get-Item $OutFile).Length / 1MB
    Write-Host ("  完了: {0:N1} MB / {1:N1} 秒" -f $mb, $sw.Elapsed.TotalSeconds)
}

function Install-Git {
    param([bool]$Portable)

    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Git は既に導入済みです: $((git --version))" -ForegroundColor Green
        return
    }

    Write-Host "`n=== Git for Windows ===" -ForegroundColor Cyan
    # 最新版のダウンロード URL を GitHub API から解決する（URL 直書きは版上げで腐るため）
    $api = 'https://api.github.com/repos/git-for-windows/git/releases/latest'
    $rel = Invoke-RestMethod -Uri $api -UseBasicParsing -Headers @{ 'User-Agent' = 'ps' }

    if ($Portable) {
        $asset = $rel.assets | Where-Object { $_.name -like 'PortableGit-*-64-bit.7z.exe' } | Select-Object -First 1
        if (-not $asset) { throw 'PortableGit のアセットが見つかりません' }
        $exe = Join-Path $env:TEMP $asset.name
        Get-File -Uri $asset.browser_download_url -OutFile $exe

        $dest = Join-Path $InstallRoot 'PortableGit'
        New-Item -ItemType Directory -Force -Path $dest | Out-Null
        Write-Host "  展開中: $dest"
        # 7-Zip 自己展開書庫。-o で展開先、-y で確認なし
        Start-Process -FilePath $exe -ArgumentList "-o`"$dest`"", '-y' -Wait -NoNewWindow

        $gitBin = Join-Path $dest 'cmd'
        # このセッションと、ユーザー環境変数の両方に通す
        $env:Path = "$gitBin;$env:Path"
        $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
        if ($userPath -notlike "*$gitBin*") {
            [Environment]::SetEnvironmentVariable('Path', "$gitBin;$userPath", 'User')
        }
        Write-Host "  PATH に追加: $gitBin（新しいコンソールから有効）" -ForegroundColor Green
    } else {
        $asset = $rel.assets | Where-Object { $_.name -like 'Git-*-64-bit.exe' } | Select-Object -First 1
        if (-not $asset) { throw 'Git インストーラのアセットが見つかりません' }
        $exe = Join-Path $env:TEMP $asset.name
        Get-File -Uri $asset.browser_download_url -OutFile $exe

        Write-Host '  サイレントインストール中（管理者権限が必要）'
        Start-Process -FilePath $exe -Wait -NoNewWindow `
            -ArgumentList '/VERYSILENT', '/NORESTART', '/NOCANCEL', '/SP-', '/CLOSEAPPLICATIONS'
        $env:Path = "$env:ProgramFiles\Git\cmd;$env:Path"
    }
    Write-Host '  Git 導入完了' -ForegroundColor Green
}

function Install-Code {
    param([bool]$NoAdminMode)

    if (Get-Command code -ErrorAction SilentlyContinue) {
        Write-Host 'VS Code は既に導入済みです' -ForegroundColor Green
        return
    }

    Write-Host "`n=== Visual Studio Code ===" -ForegroundColor Cyan
    # update.code.visualstudio.com の安定エイリアス。版を問わず常に最新へ解決される
    #   win32-x64-user    : ユーザーインストーラ（管理者権限不要）
    #   win32-x64         : システムインストーラ（管理者権限が必要）
    #   win32-x64-archive : zip（インストール不要のポータブル）
    if ($NoAdminMode) {
        $uri = 'https://update.code.visualstudio.com/latest/win32-x64-user/stable'
        $exe = Join-Path $env:TEMP 'VSCodeUserSetup.exe'
        Get-File -Uri $uri -OutFile $exe
        Write-Host '  サイレントインストール中（ユーザー領域）'
        Start-Process -FilePath $exe -Wait -NoNewWindow `
            -ArgumentList '/VERYSILENT', '/NORESTART', '/MERGETASKS=!runcode,addtopath'
        $env:Path = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin;$env:Path"
    } else {
        $uri = 'https://update.code.visualstudio.com/latest/win32-x64/stable'
        $exe = Join-Path $env:TEMP 'VSCodeSetup.exe'
        Get-File -Uri $uri -OutFile $exe
        Write-Host '  サイレントインストール中（システム全体・管理者権限が必要）'
        Start-Process -FilePath $exe -Wait -NoNewWindow `
            -ArgumentList '/VERYSILENT', '/NORESTART', '/MERGETASKS=!runcode,addtopath'
        $env:Path = "$env:ProgramFiles\Microsoft VS Code\bin;$env:Path"
    }
    Write-Host '  VS Code 導入完了（新しいコンソールから code コマンドが使えます）' -ForegroundColor Green
}

# --- 実行 --------------------------------------------------------------------
$isAdmin = Test-Admin
Write-Host ('管理者権限: ' + $(if ($isAdmin) { 'あり' } else { 'なし' }))
if (-not $isAdmin -and -not $NoAdmin) {
    Write-Host '管理者権限がないため、-NoAdmin モード（ポータブル版）に切り替えます。' -ForegroundColor Yellow
    $NoAdmin = $true
}

# 疎通確認を先に行う。落ちるならここで分かった方が早い。
# 落とす対象そのものではなく軽量な API を叩く（本体は 60〜100MB あるため）
Write-Host "`n外向き HTTPS の疎通確認..."
$probes = @{
    'update.code.visualstudio.com' = 'https://update.code.visualstudio.com/api/update/win32-x64/stable/latest'
    'api.github.com'               = 'https://api.github.com/repos/git-for-windows/git/releases/latest'
}
$reachable = $true
foreach ($name in $probes.Keys) {
    try {
        Invoke-WebRequest -Uri $probes[$name] -UseBasicParsing -TimeoutSec 15 | Out-Null
        Write-Host ("  OK   " + $name) -ForegroundColor Green
    } catch {
        Write-Host ("  NG   " + $name + " : " + $_.Exception.Message) -ForegroundColor Red
        $reachable = $false
    }
}
if (-not $reachable) {
    Write-Host ''
    Write-Host '外向き HTTPS に到達できないホストがあります。' -ForegroundColor Yellow
    Write-Host 'その場合は tools/make_transfer_bundle.py で作るクリップボード転送を使ってください。'
    Write-Host '（プロキシ環境なら $env:HTTPS_PROXY を設定してから再実行）'
    exit 1
}

if ($What -eq 'git' -or $What -eq 'all') { Install-Git -Portable:$NoAdmin }
if ($What -eq 'code' -or $What -eq 'all') { Install-Code -NoAdminMode:$NoAdmin }

Write-Host "`n完了しました。新しい PowerShell を開いてから使ってください。" -ForegroundColor Green
