@echo off
echo Building Discord Enhancements NVDA Add-on...
cd /d "%~dp0"
if exist "DiscordEnhancements.nvda-addon" del "DiscordEnhancements.nvda-addon"
powershell -Command "Compress-Archive -Path 'manifest.ini','installTasks.py','appModules','globalPlugins','doc' -DestinationPath 'DiscordEnhancements.zip' -Force"
ren "DiscordEnhancements.zip" "DiscordEnhancements.nvda-addon"
echo.
echo Done! Output: DiscordEnhancements.nvda-addon
pause
