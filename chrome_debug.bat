@echo off
:: ============================================================
::  SGPilot — Abre o Chrome com depuração remota ativa
::
::  IMPORTANTE: feche todo o Chrome antes de executar.
:: ============================================================

echo Fechando instâncias anteriores do Chrome...
taskkill /F /IM chrome.exe /T >nul 2>&1
timeout /t 2 /nobreak >nul

echo Abrindo Chrome com depuração na porta 9222...

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%APPDATA%\SGPilotChrome"

echo.
echo Chrome aberto! Acesse o SGP normalmente.
echo Depois clique em "Conectar ao Chrome" no SGPilot.
echo.
timeout /t 3 /nobreak >nul
