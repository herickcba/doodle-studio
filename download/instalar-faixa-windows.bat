@echo off
setlocal EnableExtensions
title CBA Studio - Instalador da faixa (Windows) [beta]

rem ============================================================
rem  Instalador da faixa "CBA Studio" (.ppam) no PowerPoint do
rem  Windows. Copia o add-in para a pasta AddIns do Office e o
rem  registra (AutoLoad + Path) para a aba aparecer sozinha.
rem
rem  BETA: valide com o checklist WINDOWS-TEST.md antes de
rem  distribuir ao time. Se a aba nao aparecer, use o metodo
rem  manual descrito no fim.
rem ============================================================

rem Procura o .ppam ao lado do script, depois em Downloads
set "SRC=%~dp0BG-DoodleStudio.ppam"
if not exist "%SRC%" set "SRC=%USERPROFILE%\Downloads\BG-DoodleStudio.ppam"
set "DESTDIR=%APPDATA%\Microsoft\AddIns"
set "DEST=%DESTDIR%\BG-DoodleStudio.ppam"

echo.
echo === CBA Studio - faixa do PowerPoint (Windows) ===
echo.

if not exist "%SRC%" (
  echo [ERRO] Nao encontrei o add-in "BG-DoodleStudio.ppam".
  echo        Procurei ao lado deste script e em %USERPROFILE%\Downloads
  echo.
  echo Baixe "BG-DoodleStudio.ppam" para a pasta Downloads e rode de novo.
  echo.
  pause
  exit /b 1
)

rem PowerPoint aberto? Pede para o usuario fechar (sem matar - preserva trabalho nao salvo)
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo [ATENCAO] O PowerPoint esta aberto.
  echo Salve seu trabalho, FECHE o PowerPoint e depois pressione qualquer tecla...
  pause >nul
  tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
  if not errorlevel 1 (
    echo [ERRO] O PowerPoint ainda esta aberto. Feche-o e rode este instalador de novo.
    pause
    exit /b 1
  )
)

if not exist "%DESTDIR%" mkdir "%DESTDIR%"
copy /Y "%SRC%" "%DEST%" >nul
if errorlevel 1 (
  echo [ERRO] Nao consegui copiar para "%DEST%".
  pause
  exit /b 1
)

rem Registra: HKCU\...\16.0\PowerPoint\AddIns\DoodleStudio (Office 2016/365)
set "KEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\DoodleStudio"
reg add "%KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul
if errorlevel 1 (
  echo [ERRO] Nao consegui registrar o add-in ^(reg add falhou^).
  pause
  exit /b 1
)
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
