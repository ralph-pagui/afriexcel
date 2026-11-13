@echo off
echo Installing office-addin-dev-settings...
node "%ProgramFiles%\nodejs\node_modules\npm\bin\npm-cli.js" install -g office-addin-dev-settings
echo.
echo Installation done. Now sideloading...
echo.

:: Find the installed package location
for /f "delims=" %%i in ('node -e "console.log(require('child_process').execSync('node \"%ProgramFiles%\\nodejs\\node_modules\\npm\\bin\\npm-cli.js\" root -g').toString().trim())"') do set GLOBAL_MODULES=%%i

echo Global modules: %GLOBAL_MODULES%

:: Try to sideload
node "%ProgramFiles%\nodejs\node_modules\npm\bin\npx-cli.js" office-addin-dev-settings sideload "%LOCALAPPDATA%\AssistantIAExcel\manifest.xml" --app excel

echo.
echo Done!
pause
