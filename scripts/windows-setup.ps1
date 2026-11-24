# 나라장터 MCP - Windows 최초 설치 (1_설치.bat 더블클릭으로 실행)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Find-Ngrok {
    $c = Get-Command ngrok -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $p = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\ngrok.exe"
    if (Test-Path $p) { return $p }
    return $null
}

Write-Host "1) Python 확인/설치" -ForegroundColor Cyan
if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    Write-Host ""
    Write-Host "Python 설치 완료. 이 창을 닫고 1_설치.bat 을 한 번 더 더블클릭하세요." -ForegroundColor Yellow
    exit 0
}

Write-Host "2) 프로그램 설치 (몇 분 걸립니다)" -ForegroundColor Cyan
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip -q
.\.venv\Scripts\python.exe -m pip install -r requirements.txt uvicorn starlette -q

Write-Host "3) ngrok 설치/업데이트" -ForegroundColor Cyan
$ng = Find-Ngrok
if (-not $ng) {
    winget install -e --id Ngrok.Ngrok --accept-source-agreements --accept-package-agreements
    $ng = Find-Ngrok
}
if ($ng) { try { & $ng update } catch { } } else { Write-Host "ngrok을 찾지 못했습니다. 창을 닫고 1_설치.bat 을 다시 실행하세요." -ForegroundColor Red; exit 1 }

Write-Host "4) 설정값 입력 (붙여넣기는 마우스 오른쪽 클릭)" -ForegroundColor Cyan
$key = ((Read-Host "공공데이터포털 일반 인증키") -replace '\s','')
$domain = ((Read-Host "ngrok 도메인 (예: xxxx.ngrok-free.dev)") -replace '\s','') -replace '^https?://','' -replace '/.*$',''
$ngTok = ((Read-Host "ngrok Authtoken") -replace '\s','')
& $ng config add-authtoken $ngTok

$bytes = New-Object byte[] 24; [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
$token = ([Convert]::ToBase64String($bytes)) -replace '[+/=]',''

$enc = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $root ".env.local"), "NARAMARKET_SERVICE_KEY=$key`r`nMCP_ACCESS_TOKEN=$token`r`nNGROK_DOMAIN=$domain`r`n", $enc)

$url = "https://$domain/mcp?token=$token"
Set-Content -Path (Join-Path $root "커넥터주소.txt") -Value $url -Encoding UTF8
Write-Host ""
Write-Host "설치 완료!" -ForegroundColor Green
Write-Host "Claude 커넥터에 넣을 주소 (커넥터주소.txt 에도 저장됨):" -ForegroundColor Green
Write-Host $url -ForegroundColor Yellow
Write-Host "이제 이 창을 닫고 2_서버켜기.bat 을 더블클릭하세요." -ForegroundColor Green
