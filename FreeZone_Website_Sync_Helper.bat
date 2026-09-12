@echo off
title FreeZone Website Sync Helper
color 0b

echo ======================================================
echo           FREEZONE WEBSITE SYNC HELPER
echo ======================================================
echo.

:MENU
echo [1] START WORKING (Pull Latest from Cloud)
echo [2] FINISH & PUBLISH (Push Changes to Cloud)
echo [3] EXIT
echo.
set /p choice="Please select an option (1-3): "

if "%choice%"=="1" goto PULL
if "%choice%"=="2" goto PUSH
if "%choice%"=="3" goto EXIT
goto MENU

:PULL
echo.
echo [STEP 1/1] Fetching latest updates from GitHub...
git pull origin main
echo.
echo ======================================================
echo SUCCESS: Your local files are now up to date!
echo You can now start editing your website.
echo ======================================================
pause
goto MENU

:PUSH
echo.
echo [STEP 1/3] Packing changes...
git add .
echo.
echo [STEP 2/3] Adding timestamped label...
set dt=%date% %time%
git commit -m "Automated Sync Update: %dt%"
echo.
echo [STEP 3/3] Uploading to GitHub & Triggering Netlify...
git push origin master:main -f
echo.
echo ======================================================
echo SUCCESS: Your website is being updated globally!
echo Please check www.freezoneconcept.com in 30 seconds.
echo ======================================================
pause
goto MENU

:EXIT
exit
