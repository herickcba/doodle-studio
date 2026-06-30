@echo off
setlocal EnableExtensions
title CBA Studio - Instalador da faixa (Windows) [beta]

rem ============================================================
rem  Instalador da faixa "CBA Studio" (.ppam) no PowerPoint do
rem  Windows. Copia o add-in para a pasta AddIns do Office e o
rem  registra (AutoLoad + Path) para a aba aparecer sozinha.
rem
rem  BETA: nao foi testado pelo autor em Windows. Se a aba nao
rem  aparecer, use o metodo manual descrito no fim.
rem ============================================================

set "SRC=%USERPROFILE%\Downloads\BG-DoodleStudio.ppam"
set "DESTDIR=%APPDATA%\Microsoft\AddIns"
set "DEST=%DESTDIR%\BG-DoodleStudio.ppam"

echo.
echo === CBA Studio - faixa do PowerPoint (Windows) ===
echo.

if not exist "%SRC%" (
  echo [ERRO] Nao encontrei o add-in em:
  echo        %SRC%
  echo.
  echo Baixe "BG-DoodleStudio.ppam" para a pasta Downloads e rode de novo.
  echo.
  pause
  exit /b 1
)

if not exist "%DESTDIR%" mkdir "%DESTDIR%"
copy /Y "%SRC%" "%DEST%" >nul
if errorlevel 1 (
  echo [ERRO] Nao consegui copiar para "%DEST%".
  pause
  exit /b 1
)

rem Fecha o PowerPoint para recarregar o add-in na proxima abertura
taskkill /IM POWERPNT.EXE /F >nul 2>&1

rem Registra: HKCU\...\16.0\PowerPoint\AddIns\DoodleStudio
set "KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\DoodleStudio"
reg add "%KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul
reg add "%KEY%" /v Path /t REG_SZ /d "%DEST%" /f >nul

echo [OK] Instalado em:
echo      %DEST%
echo.
echo Abra o PowerPoint. Se pedir, clique "Habilitar Macros".
echo A aba "CBA Studio" deve aparecer na faixa.
echo.
echo --- Se a aba NAO aparecer, metodo manual ---
echo   Arquivo  ^>  Opcoes  ^>  Suplementos
echo   Em "Gerenciar", escolha "Suplementos do PowerPoint"  ^>  Ir...
echo   Clique "Adicionar Novo..." e selecione:
echo      %DEST%
echo.
pause
