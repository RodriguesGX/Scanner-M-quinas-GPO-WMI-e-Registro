@echo off
:: ===================================================
:: REPARO RAPIDO DO WMI - Windows 7
:: Versao: 1.0
:: Executa antes do Scanner para corrigir 0x80041003
:: ===================================================
setlocal enabledelayedexpansion

set "LOG=C:\ProgramData\ScannerMaquinas\Logs\wmi_repair.log"
if not exist "C:\ProgramData\ScannerMaquinas\Logs" mkdir "C:\ProgramData\ScannerMaquinas\Logs"

echo %date% %time% - Iniciando reparo do WMI >> "%LOG%"

:: Verifica se o WMI está responsivo
wmic os get Caption >nul 2>&1
if !errorlevel! equ 0 (
    echo %date% %time% - WMI OK, nenhum reparo necessario >> "%LOG%"
    exit /b 0
)

echo %date% %time% - WMI nao respondeu, iniciando reparo... >> "%LOG%"

:: Método 1: Recompilar MOFs (rápido, resolve 70% dos casos)
echo %date% %time% - Metodo 1: Recompilando MOFs >> "%LOG%"
cd /d %windir%\system32\wbem
for %%f in (*.mof) do (
    mofcomp "%%f" >nul 2>&1
)

:: Testa novamente
wmic os get Caption >nul 2>&1
if !errorlevel! equ 0 (
    echo %date% %time% - WMI recuperado apos Metodo 1 >> "%LOG%"
    exit /b 0
)

:: Método 2: Recriar repositório (mais lento, resolve 95% dos casos)
echo %date% %time% - Metodo 2: Recriando repositorio >> "%LOG%"
net stop winmgmt /y >nul 2>&1
timeout /t 3 /nobreak >nul
if exist "%windir%\system32\wbem\Repository.old" rmdir /s /q "%windir%\system32\wbem\Repository.old" >nul 2>&1
ren "%windir%\system32\wbem\Repository" "Repository.old" >nul 2>&1
net start winmgmt >nul 2>&1
timeout /t 5 /nobreak >nul

:: Recompila MOFs novamente
cd /d %windir%\system32\wbem
for %%f in (*.mof) do (
    mofcomp "%%f" >nul 2>&1
)

:: Testa novamente
wmic os get Caption >nul 2>&1
if !errorlevel! equ 0 (
    echo %date% %time% - WMI recuperado apos Metodo 2 >> "%LOG%"
    exit /b 0
)

:: Método 3: Salvage (último recurso)
echo %date% %time% - Metodo 3: Salvage repository >> "%LOG%"
net stop winmgmt /y >nul 2>&1
timeout /t 3 /nobreak >nul
cd /d %windir%\system32\wbem
winmgmt /salvagerepository >nul 2>&1
net start winmgmt >nul 2>&1
timeout /t 5 /nobreak >nul

:: Testa final
wmic os get Caption >nul 2>&1
if !errorlevel! equ 0 (
    echo %date% %time% - WMI recuperado apos Metodo 3 >> "%LOG%"
) else (
    echo %date% %time% - FALHA: WMI nao pode ser reparado >> "%LOG%"
)

echo %date% %time% - Reparo finalizado >> "%LOG%"
exit /b 0