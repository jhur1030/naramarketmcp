# 나라장터 MCP - 서버 켜기 (창을 닫으면 서버도 꺼집니다)
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# 최신 코드 자동 받기 (실패해도 기존 코드로 계속 실행)
try {
    $zip = Join-Path $env:TEMP "naramarket-latest.zip"
    $tmp = Join-Path $env:TEMP "naramarket-latest"
    Invoke-WebRequest "https://github.com/jhur1030/naramarketmcp/archive/refs/heads/main.zip" -OutFile $zip -UseBasicParsing
    if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
    Expand-Archive $zip $tmp -Force
    $src = Join-Path $tmp "naramarketmcp-main"
    robocopy (Join-Path $src "src") (Join-Path $root "src") /E /NFL /NDL /NJH /NJS /NP | Out-Null
    robocopy (Join-Path $src "scripts") (Join-Path $root "scripts") /E /NFL /NDL /NJH /NJS /NP | Out-Null
    Write-Host "최신 코드로 업데이트했습니다." -ForegroundColor Cyan
} catch { Write-Host "업데이트 건너뜀: $($_.Exception.Message)" -ForegroundColor DarkYellow }
Get-Content .env.local | ForEach-Object {
    if ($_ -match '^\s*([^=]+)=(.*)$') { [Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim()) }
}
$env:FASTMCP_TRANSPORT = "http"
$env:FASTMCP_HOST = "127.0.0.1"
$env:PORT = "8081"

$server = Start-Process -PassThru -NoNewWindow .\.venv\Scripts\python.exe -ArgumentList "-m","src.main"
Start-Sleep -Seconds 5
if ($env:NGROK_DOMAIN) {
    Write-Host "휴대폰/Claude 커넥터 주소:" -ForegroundColor Green
    Write-Host "https://$($env:NGROK_DOMAIN)/mcp?token=$($env:MCP_ACCESS_TOKEN)" -ForegroundColor Yellow
    $ng = (Get-Command ngrok -ErrorAction SilentlyContinue).Source
    if (-not $ng) { $ng = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\ngrok.exe" }
    & $ng http --domain=$env:NGROK_DOMAIN 8081
} else {
    Write-Host "PC 전용 모드로 실행 중 (http://127.0.0.1:8081/mcp). 끄려면 이 창을 닫으세요." -ForegroundColor Green
    Wait-Process -Id $server.Id
}
Stop-Process -Id $server.Id -ErrorAction SilentlyContinue
