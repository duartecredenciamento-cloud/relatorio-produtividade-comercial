@echo off
setlocal
cd /d "%~dp0"
echo Instalando o Duarte - Produtividade Comercial...
where py >nul 2>nul
if errorlevel 1 (
  echo Python nao foi encontrado. Instale Python 3.10 ou mais recente e marque Add Python to PATH.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
if errorlevel 1 goto :erro
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto :erro
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto :erro
echo.
echo Instalacao concluida. Agora abra iniciar.bat.
pause
exit /b 0
:erro
echo A instalacao nao terminou. Confira sua conexao e a instalacao do Python.
pause
exit /b 1
