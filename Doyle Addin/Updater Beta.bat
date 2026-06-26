@echo off
SETLOCAL EnableDelayedExpansion

REM --- Optional: Inventor process ID to wait for before deleting ---
SET "INVENTOR_PID=%~1"

REM --- Wait for Inventor to fully exit before touching its files ---
IF NOT "%INVENTOR_PID%"=="" (
    ECHO Waiting for Inventor process %INVENTOR_PID% to exit...
    :WAIT_FOR_EXIT
    timeout /t 2 /nobreak >NUL
    tasklist /FI "PID eq %INVENTOR_PID%" 2>NUL | find /I "Inventor.exe" >NUL
    IF NOT ERRORLEVEL 1 GOTO WAIT_FOR_EXIT
    ECHO Inventor process exited. Proceeding with update.
)

timeout /t 3 /nobreak
REM --- Configuration ---
SET GITHUB_USER=LeviErskine
SET GITHUB_REPO=Doyle-AddIn-C-
SET ADDIN_NAME=DoyleAddin

REM --- Instructions ---
REM This script is designed to be placed directly inside the Add-in's installation folder
REM and run from there. It will automatically detect its location and update the contents.

REM --- Define target directory (automatically detected from this script's location) ---
SET "TARGET_ADDINS_PATH=%~dp0"
IF "%TARGET_ADDINS_PATH:~-1%"=="\" SET "TARGET_ADDINS_PATH=%TARGET_ADDINS_PATH:~0,-1%"

REM --- Get the name of this script to ensure it is not deleted during cleanup ---
SET "SCRIPT_FILENAME=%~nx0"

ECHO.
ECHO Updating %ADDIN_NAME% Add-in (BETA channel)...
ECHO Target directory: "%TARGET_ADDINS_PATH%"
ECHO Installer script: "!SCRIPT_FILENAME!"
ECHO.

REM --- Clean up existing installation (deletes all files/folders except this script) ---
ECHO Removing existing add-in files...

FOR /F "delims=" %%i IN ('dir /b /a-d "%TARGET_ADDINS_PATH%"') DO (
    IF /I NOT "%%i" == "!SCRIPT_FILENAME!" (
        IF /I NOT "%%i" == "Updater.bat" (
            ECHO  - Deleting file: %%i
            DEL /F /Q "%TARGET_ADDINS_PATH%\%%i"
        )
    )
)

FOR /D %%i IN ("%TARGET_ADDINS_PATH%\*") DO (
    ECHO  - Deleting folder: %%i
    RMDIR /S /Q "%%i"
)
ECHO Cleanup of old files complete.
ECHO.

REM --- Find the latest pre-release via GitHub API ---
ECHO Looking up latest pre-release...
SET API_URL=https://api.github.com/repos/%GITHUB_USER%/%GITHUB_REPO%/releases

powershell -Command "try { $r = Invoke-RestMethod -Uri '%API_URL%' -UseBasicParsing | Where-Object { $_.prerelease -eq $true } | Select-Object -First 1; if ($r) { $a = $r.assets | Where-Object { $_.name -eq '%ADDIN_NAME%.zip' } | Select-Object -First 1; if ($a) { Write-Output $a.browser_download_url } else { Write-Output 'ASSET_NOT_FOUND' } } else { Write-Output 'NO_PRERELEASE' } } catch { Write-Output 'API_ERROR' }" > "%TEMP%\gh_download_url.txt"
SET /P DOWNLOAD_URL=<"%TEMP%\gh_download_url.txt"

IF "%DOWNLOAD_URL%"=="NO_PRERELEASE" (
    ECHO ERROR: No pre-release found for this repository.
    PAUSE
    GOTO :EOF
)
IF "%DOWNLOAD_URL%"=="ASSET_NOT_FOUND" (
    ECHO ERROR: Pre-release found but %ADDIN_NAME%.zip asset is missing.
    PAUSE
    GOTO :EOF
)
IF "%DOWNLOAD_URL%"=="API_ERROR" (
    ECHO ERROR: Failed to query GitHub API. Check your internet connection.
    PAUSE
    GOTO :EOF
)

SET ZIP_FILENAME=%ADDIN_NAME%.zip
SET TARGET_ZIP_PATH="%TARGET_ADDINS_PATH%\%ZIP_FILENAME%"

ECHO Downloading latest pre-release from %DOWNLOAD_URL%...
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TARGET_ZIP_PATH%' -UseBasicParsing"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to download the zip file.
    PAUSE
    GOTO :EOF
)

ECHO Download complete. Extracting new files...
powershell -Command "Expand-Archive -LiteralPath '%TARGET_ZIP_PATH%' -DestinationPath '%TARGET_ADDINS_PATH%' -Force"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to extract the zip file.
    PAUSE
    GOTO :CLEANUP_ZIP
)

ECHO.
ECHO %ADDIN_NAME% Add-in (BETA) updated successfully!
ECHO It should load next time Inventor starts.
ECHO.

:CLEANUP_ZIP
IF EXIST "%TARGET_ZIP_PATH%" (
    ECHO Deleting temporary zip file...
    DEL /Q "%TARGET_ZIP_PATH%"
)

ECHO Update process finished.

ENDLOCAL
GOTO :EOF
