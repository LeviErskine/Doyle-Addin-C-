@echo off
SETLOCAL EnableDelayedExpansion

REM --- Configuration ---
SET GITHUB_USER=Bmassner
SET GITHUB_REPO=Doyle-AddIn
SET ADDIN_NAME=DoyleAddin
REM This should be the base name of your add-in files (e.g., DoyleAddin.dll, DoyleAddin.addin)

REM --- Instructions ---
REM This script is designed to be placed directly inside the Add-in's installation folder
REM (e.g., C:\ProgramData\Autodesk\Inventor Addins\DoyleAddin) and run from there.
REM It will automatically detect its location and update the contents of that folder.

REM --- Define target directory (automatically detected from this script's location) ---
SET "TARGET_ADDINS_PATH=%~dp0"
REM Clean up the trailing backslash from the path for consistency
IF "%TARGET_ADDINS_PATH:~-1%"=="\" SET "TARGET_ADDINS_PATH=%TARGET_ADDINS_PATH:~0,-1%"

REM --- Get the name of this script to ensure it is not deleted during cleanup ---
SET "SCRIPT_FILENAME=%~nx0"

ECHO.
ECHO Updating %ADDIN_NAME% Add-in...
ECHO Target directory: "%TARGET_ADDINS_PATH%"
ECHO Installer script: "!SCRIPT_FILENAME!"
ECHO.

REM --- Clean up existing installation (deletes all files/folders except this script) ---
ECHO Removing existing add-in files...

REM Delete all files in the current directory, EXCEPT for this running script.
FOR /F "delims=" %%i IN ('dir /b /a-d "%TARGET_ADDINS_PATH%"') DO (
    IF /I NOT "%%i" == "!SCRIPT_FILENAME!" (
        ECHO  - Deleting file: %%i
        DEL /F /Q "%TARGET_ADDINS_PATH%\%%i"
    )
)

REM Delete all subdirectories in the current directory.
FOR /D %%i IN ("%TARGET_ADDINS_PATH%\*") DO (
    ECHO  - Deleting folder: %%i
    RMDIR /S /Q "%%i"
)
ECHO Cleanup of old files complete.
ECHO.

REM --- Download the latest release ZIP from GitHub directly to the target folder ---
SET ZIP_FILENAME=%ADDIN_NAME%.zip
SET DOWNLOAD_URL="https://github.com/%GITHUB_USER%/%GITHUB_REPO%/releases/latest/download/%ZIP_FILENAME%"
SET TARGET_ZIP_PATH="%TARGET_ADDINS_PATH%\%ZIP_FILENAME%"

ECHO Downloading latest release from %DOWNLOAD_URL%...
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TARGET_ZIP_PATH%' -UseBasicParsing"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to download the zip file. Check your internet connection or the GitHub URL.
    PAUSE
    GOTO :EOF
)

ECHO Download complete. Extracting new files...

REM --- Extract the ZIP file directly into the target directory ---
REM The -DestinationPath will extract the contents of the ZIP to TARGET_ADDINS_PATH.
REM IMPORTANT: Ensure your ZIP file contains the DLL and .addin at its root.
powershell -Command "Expand-Archive -LiteralPath '%TARGET_ZIP_PATH%' -DestinationPath '%TARGET_ADDINS_PATH%' -Force"
IF %ERRORLEVEL% NEQ 0 (
    ECHO ERROR: Failed to extract the zip file.
    PAUSE
    GOTO :CLEANUP_ZIP
)

ECHO.
ECHO %ADDIN_NAME% Add-in updated successfully!
ECHO It should load next time Inventor starts.
ECHO.

:CLEANUP_ZIP
REM Delete the downloaded zip file after extraction
IF EXIST "%TARGET_ZIP_PATH%" (
    ECHO Deleting temporary zip file...
    DEL /Q %TARGET_ZIP_PATH%
)

ECHO Update process finished.
ECHO This installer will now delete itself. Press any key to finish.
PAUSE

REM --- Self-destruct mechanism ---
REM Starts a new, non-blocking cmd process that waits 1 second then deletes this script file.
REM This allows the main script to exit cleanly before deletion occurs.
start /b "" cmd /c "ping 127.0.0.1 -n 2 > nul & del /F /Q "%~f0""

ENDLOCAL
GOTO :EOF