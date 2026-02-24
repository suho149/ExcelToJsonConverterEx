@echo off
setlocal

cd /d "%~dp0\.."

echo [1/4] Building fat jar...
call mvnw.cmd -q -DskipTests clean package
if errorlevel 1 (
    echo Maven package failed.
    exit /b 1
)

set "SHADED_JAR="
for %%F in (target\*-all.jar) do (
    set "SHADED_JAR=%%~nxF"
    goto jar_found
)

echo Shaded jar not found in target folder.
exit /b 1

:jar_found
where jpackage >nul 2>nul
if errorlevel 1 (
    echo jpackage command not found. Please install JDK 17+ and add it to PATH.
    exit /b 1
)

set "APP_NAME=ExcelToJsonConverter"
set "DIST_DIR=dist"

if exist "%DIST_DIR%\%APP_NAME%" (
    echo [0/4] Removing previous app-image...
    del /f /q "%DIST_DIR%\%APP_NAME%\excel\~$*.xlsx" >nul 2>nul
    rmdir /s /q "%DIST_DIR%\%APP_NAME%" >nul 2>nul
    if exist "%DIST_DIR%\%APP_NAME%" (
        echo Failed to remove "%DIST_DIR%\%APP_NAME%".
        echo Close these processes and run again:
        echo   - EXCEL.EXE
        echo   - %APP_NAME%.exe
        echo Also close any File Explorer window opened in the dist folder.
        exit /b 1
    )
)

echo [2/4] Packaging Windows app-image...
jpackage --type app-image ^
  --name "%APP_NAME%" ^
  --input target ^
  --main-jar "%SHADED_JAR%" ^
  --main-class demo.tojson.ExcelToJsonConverter ^
  --dest "%DIST_DIR%" ^
  --win-console

if errorlevel 1 (
    echo jpackage failed.
    exit /b 1
)

echo [3/4] Copying default Excel file...
set "RESOURCE_DIR=%DIST_DIR%\%APP_NAME%\excel"
if not exist "%RESOURCE_DIR%" mkdir "%RESOURCE_DIR%"
copy /Y "src\main\resources\excel\exceldata.xlsx" "%RESOURCE_DIR%\exceldata.xlsx" >nul

echo [4/4] Copying README...
copy /Y "scripts\README.txt" "%DIST_DIR%\%APP_NAME%\README.txt" >nul

echo.
echo Build completed.
echo Run this file:
echo   %DIST_DIR%\%APP_NAME%\%APP_NAME%.exe
echo.
echo Default input file:
echo   %DIST_DIR%\%APP_NAME%\excel\exceldata.xlsx

exit /b 0
