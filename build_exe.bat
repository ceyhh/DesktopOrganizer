@echo off
setlocal

set APP_NAME=DesktopOrganizer
set SCRIPT_NAME=desktop_extension_organizer.py
set RELEASE_DIR=release

python --version >nul 2>nul
if errorlevel 1 (
  echo Python bulunamadi. Lutfen once Python kur.
  pause
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install pyinstaller

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"

pyinstaller --noconfirm --clean --onefile --noconsole --name "%APP_NAME%" "%SCRIPT_NAME%"

if not exist dist\%APP_NAME%.exe (
  echo Build basarisiz oldu, exe bulunamadi.
  pause
  exit /b 1
)

if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%"

copy /y dist\%APP_NAME%.exe "%RELEASE_DIR%\%APP_NAME%.exe" >nul
copy /y README.md "%RELEASE_DIR%\README.txt" >nul

echo.
echo EXE hazir: dist\%APP_NAME%.exe
echo Release hazir: %RELEASE_DIR%\%APP_NAME%.exe
echo Release notlari: %RELEASE_DIR%\README.txt
pause
