@echo off
setlocal

set "SRC=D:\Work\Tools\PremiereProjectCollector"
set "DEST=%APPDATA%\Adobe\CEP\extensions\PremiereProjectCollector"
set "PARENT=%APPDATA%\Adobe\CEP\extensions"
set "RC=0"

title Project Collector Install

echo Source: %SRC%
echo Destination: %DEST%
echo.

if not exist "%SRC%" (
    echo [ERROR] Source folder not found.
    echo.
    pause
    exit /b 1
)

if not exist "%PARENT%" (
    echo Creating CEP extensions parent folder...
    mkdir "%PARENT%"
    if errorlevel 1 (
        echo [ERROR] Could not create CEP extensions parent folder.
        echo.
        pause
        exit /b 1
    )
)

if exist "%DEST%" (
    echo Removing older installed Project Collector folder...
    rmdir /s /q "%DEST%"
    if exist "%DEST%" (
        echo [ERROR] Could not remove the existing installed folder.
        echo Close Premiere Pro and try again.
        echo.
        pause
        exit /b 1
    )
)

echo Copying fresh Project Collector files...
echo.
robocopy "%SRC%" "%DEST%" /MIR /XD .git /XF deploy_extension.bat install_extension.bat
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    echo.
    echo [ERROR] Install failed with robocopy exit code %RC%.
    echo The extension may not have been copied correctly.
    echo.
    pause
    exit /b %RC%
)

echo.
echo [DONE] Project Collector installed successfully.
echo Restart Premiere Pro if it was already open.
echo.
pause
exit /b 0
