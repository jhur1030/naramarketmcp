# 나라장터 MCP - 서버 켜기 (창을 닫으면 서버도 꺼집니다)
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
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
    ngrok http --url=$env:NGROK_DOMAIN 8081
} else {
    Write-Host "PC 전용 모드로 실행 중 (http://127.0.0.1:8081/mcp). 끄려면 이 창을 닫으세요." -ForegroundColor Green
    Wait-Process -Id $server.Id
}
Stop-Process -Id $server.Id -ErrorAction SilentlyContinue
