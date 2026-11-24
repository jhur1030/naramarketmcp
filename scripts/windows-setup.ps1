# 나라장터 MCP - Windows 최초 설치 스크립트
# 사용법: PowerShell에서  powershell -ExecutionPolicy Bypass -File scripts\windows-setup.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Host "1) Python 3.12 확인/설치" -ForegroundColor Cyan
if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    Write-Host "Python 설치 완료. 이 창을 닫고 1_설치.bat 을 한 번 더 더블클릭하세요." -ForegroundColor Yellow
    exit 0
}

Write-Host "2) 가상환경 + 패키지 설치" -ForegroundColor Cyan
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt uvicorn starlette

Write-Host "3) ngrok 설치 (휴대폰에서 쓰기 위한 외부 연결)" -ForegroundColor Cyan
if (-not (Get-Command ngrok -ErrorAction SilentlyContinue)) {
    winget install -e --id Ngrok.Ngrok --accept-source-agreements --accept-package-agreements
}

Write-Host "4) 설정값 입력" -ForegroundColor Cyan
$key   = (Read-Host "공공데이터포털 일반 인증키").Trim()
$token = (Read-Host "MCP_ACCESS_TOKEN (커넥터 주소 token= 뒤의 값)").Trim()
$domain = (Read-Host "ngrok 고정 도메인 (예: abc-def.ngrok-free.app, PC에서만 쓰면 Enter)").Trim()
if ($domain) {
    $ngTok = (Read-Host "ngrok Authtoken (ngrok 대시보드 > Your Authtoken)").Trim()
    if ($ngTok) { ngrok config add-authtoken $ngTok }
}
@"
NARAMARKET_SERVICE_KEY=$key
MCP_ACCESS_TOKEN=$token
NGROK_DOMAIN=$domain
"@ | Set-Content -Encoding UTF8 .env.local

Write-Host "설치 끝! 이제 2_서버켜기.bat 을 더블클릭하세요." -ForegroundColor Green
