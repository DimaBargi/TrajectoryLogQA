@echo off
chcp 65001 >nul
REM ============================================================
REM  בניית TrajectoryLogQA.exe מהקוד.
REM  להריץ על מחשב Windows 64-ביט עם Python 3.9+ מותקן.
REM  תוצאה:  dist\TrajectoryLogQA.exe
REM ============================================================
setlocal
cd /d "%~dp0"

echo.
echo [1/2] מתקין תלויות (numpy + openpyxl + pyinstaller)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :err

echo.
echo [2/2] בונה את התוכנה עם PyInstaller...
pyinstaller --onefile --windowed --clean --name TrajectoryLogQA ^
  --collect-all numpy --collect-all matplotlib ^
  trj_gui.py
if errorlevel 1 goto :err

echo.
echo ============================================================
echo   הצלחה!  התוכנה:  dist\TrajectoryLogQA.exe
echo   לחיצה כפולה עליה מפעילה את התוכנה.
echo   ליצירת מתקין: פתח installer.iss ב-Inno Setup (F9).
echo ============================================================
pause
exit /b 0

:err
echo.
echo [שגיאה] הבנייה נכשלה. ראה הודעות למעלה.
pause
exit /b 1
