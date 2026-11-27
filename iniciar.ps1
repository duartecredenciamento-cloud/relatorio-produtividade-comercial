$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$matrices = if ($env:MATRICES_DIR) { $env:MATRICES_DIR } else { Read-Host "Cole o caminho da pasta local de matrizes" }
if (-not (Test-Path -LiteralPath $matrices -PathType Container)) {
  throw "A pasta de matrizes não está acessível. Configure MATRICES_DIR ou informe um caminho válido."
}
if (-not (Test-Path -LiteralPath $matrices -PathType Container)) { throw "A pasta de matrizes não está acessível." }
$roster = if ($env:RESPONSIBLES_FILE) { $env:RESPONSIBLES_FILE } else { Join-Path $PSScriptRoot "dados/clientes_credenciamento_responsaveis.xlsx" }
if (-not (Test-Path -LiteralPath $roster -PathType Leaf)) {
  Write-Host "Coloque o cadastro de responsáveis em: $roster" -ForegroundColor Yellow
  $roster = Read-Host "Cole o caminho do arquivo XLSX de responsáveis"
}
if (-not (Test-Path -LiteralPath $roster -PathType Leaf)) { throw "O cadastro de responsáveis não está acessível." }
$user = Read-Host "Usuário do painel"
$securePassword = Read-Host "Senha do painel" -AsSecureString
$password = [System.Net.NetworkCredential]::new("", $securePassword).Password
$env:MATRICES_DIR = $matrices
$env:RESPONSIBLES_FILE = $roster
$env:REQUIRE_AUTH = "true"
$env:DASHBOARD_USER = $user
$env:DASHBOARD_PASSWORD = $password
try {
  if (Test-Path ".venv/Scripts/python.exe") { & ".venv/Scripts/python.exe" "app_server.py" }
  else { py -3 "app_server.py" }
} finally {
  Remove-Item Env:DASHBOARD_USER,Env:DASHBOARD_PASSWORD -ErrorAction SilentlyContinue
  Remove-Item Env:MATRICES_DIR,Env:RESPONSIBLES_FILE,Env:REQUIRE_AUTH -ErrorAction SilentlyContinue
  $password = $null
  $securePassword.Dispose()
}
