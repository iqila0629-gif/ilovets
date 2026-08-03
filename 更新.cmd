@echo off
cd /d "%~dp0"
git add -A
git commit -m "update"
if errorlevel 1 (
  echo No changes to commit.
) else (
  echo Committed, pushing...
)
git push
echo.
echo Update pushed. GitHub Pages refreshes in about 1-2 minutes.
pause
