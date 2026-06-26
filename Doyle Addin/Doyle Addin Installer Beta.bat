@echo off
SETLOCAL EnableDelayedExpansion

REM --- Configuration ---
SET GITHUB_USER=LeviErskine
SET GITHUB_REPO=Doyle-Addin-C-
SET ADDIN_NAME=DoyleAddin

REM --- Define target directory ---
SET TARGET_ADDINS_PATH=C:\ProgramData\Autodesk\Inventor Addins\%ADDIN_NAME%

ECHO.
ECHO Installing %ADDIN_NAME% Add-in (BETA channel)...
ECHO Target directory: "%TARGET_ADDINS_PATH%"
ECHO.

REM --- Clean up existing installation ---
IF EXIST "%TARGET_ADDINS_PATH%" (
    ECHO Removing existing installation at "%TARGET_ADDINS_PATH%"...
    RMDIR /S /Q "%TARGET_ADDINS_PATH%"
    IF %ERRORLEVEL% NEQ 0 (
        ECHO ERROR: Failed to remove existing directory. Ensure Inventor is closed.
        PAUSE
        GOTO :EOF
    )
)

REM --- Create target directory ---
ECHO Creating directory: "%TARGET_ADDINS_PATH%"
MD "%TARGET_ADDINS_PATH%"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to create target directory.
    PAUSE
    GOTO :EOF
)

REM --- Find the latest pre-release via GitHub API ---
ECHO Looking up latest pre-release...
SET API_URL=https://api.github.com/repos/%GITHUB_USER%/%GITHUB_REPO%/releases
SET API_RESPONSE=%TEMP%\gh_releases.json

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

ECHO Downloading pre-release from %DOWNLOAD_URL%...
SET ZIP_FILENAME=%ADDIN_NAME%.zip
SET TARGET_ZIP_PATH="%TARGET_ADDINS_PATH%\%ZIP_FILENAME%"
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TARGET_ZIP_PATH%' -UseBasicParsing"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to download the zip file.
    IF EXIST "%TARGET_ZIP_PATH%" DEL "%TARGET_ZIP_PATH%"
    PAUSE
    GOTO :CLEANUP
)

ECHO Download complete. Extracting files...
powershell -Command "Expand-Archive -LiteralPath '%TARGET_ZIP_PATH%' -DestinationPath '%TARGET_ADDINS_PATH%' -Force"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to extract the zip file.
    PAUSE
    GOTO :CLEANUP
)

ECHO.
ECHO %ADDIN_NAME% Add-in (BETA) installed successfully!
ECHO It should load next time Inventor starts.
ECHO.

:CLEANUP
IF EXIST "%TARGET_ZIP_PATH%" (
    ECHO Deleting downloaded zip file...
    DEL %TARGET_ZIP_PATH%
)
ECHO Cleanup complete.
PAUSE
ENDLOCAL
