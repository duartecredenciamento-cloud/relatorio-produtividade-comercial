@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  .venv\Scripts\python.exe app_server.py
) else (
  where py >nul 2>nul
  if errorlevel 1 (
    echo Python nao foi encontrado. Execute instalar.bat primeiro.
    pause
    exit /b 1
  )
  py -3 app_server.py
)
if errorlevel 1 (
  echo.
  echo O painel foi encerrado com erro. Confira a configuracao e appdata\atualizacoes.log.
  pause
)
