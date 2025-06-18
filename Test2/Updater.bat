@echo off
SETLOCAL EnableDelayedExpansion

REM --- Configuration ---
SET GITHUB_USER=Bmassner
SET GITHUB_REPO=Doyle-AddIn
SET ADDIN_NAME=DoyleAddin
REM This should be the base name of your add-in files (e.g., DoyleAddin.dll, DoyleAddin.addin)

REM --- Define target directory ---
SET TARGET_ADDINS_PATH=C:\ProgramData\Autodesk\Inventor Addins\%ADDIN_NAME%
REM Note: This now includes a subfolder for your addin to keep things clean and isolated.

REM --- Check for Administrator Privileges ---

ECHO.
ECHO Installing %ADDIN_NAME% Add-in...
ECHO Target directory: "%TARGET_ADDINS_PATH%"
ECHO.

REM --- Clean up existing installation (optional but good for fresh installs/reinstalls) ---
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

REM --- Download the latest release ZIP from GitHub directly to the target folder ---
SET ZIP_FILENAME=%ADDIN_NAME%.zip
SET DOWNLOAD_URL="https://github.com/%GITHUB_USER%/%GITHUB_REPO%/releases/download/Pre/%ZIP_FILENAME%"
SET TARGET_ZIP_PATH="%TARGET_ADDINS_PATH%\%ZIP_FILENAME%"

ECHO Downloading latest release from %DOWNLOAD_URL% to "%TARGET_ZIP_PATH%"...
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TARGET_ZIP_PATH%' -UseBasicParsing"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to download the zip file. Check your internet connection or the GitHub URL.
    REM Cleanup the partially downloaded zip if it exists
    IF EXIST "%TARGET_ZIP_PATH%" DEL "%TARGET_ZIP_PATH%"
    PAUSE
    GOTO :CLEANUP
)

ECHO Download complete. Extracting files from "%TARGET_ZIP_PATH%"...

REM --- Extract the ZIP file directly into the target directory ---
REM The -DestinationPath will extract the contents of the ZIP to TARGET_ADDINS_PATH.
REM IMPORTANT: Ensure your ZIP file contains the DLL and .addin at its root,
REM NOT inside another subfolder like "YourAddin-1.0.0/" inside the zip.
powershell -Command "Expand-Archive -LiteralPath '%TARGET_ZIP_PATH%' -DestinationPath '%TARGET_ADDINS_PATH%' -Force"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to extract the zip file to the target directory.
    PAUSE
    GOTO :CLEANUP
)

ECHO.
ECHO %ADDIN_NAME% Add-in installed successfully!
ECHO It should load next time Inventor starts.
ECHO.

:CLEANUP
REM Delete the downloaded zip file after extraction
IF EXIST "%TARGET_ZIP_PATH%" (
    ECHO Deleting downloaded zip file: "%TARGET_ZIP_PATH%"
    DEL %TARGET_ZIP_PATH%
)
ECHO Cleanup complete.
PAUSE
ENDLOCAL