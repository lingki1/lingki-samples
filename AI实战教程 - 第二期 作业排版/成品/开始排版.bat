@echo off
chcp 65001 >nul
pushd "%~dp0"
title 作业排版

if not exist ".venv\Scripts\python.exe" goto no_environment

set "PYTHONPATH=%~dp0src"
set "PYTHONUTF8=1"
set "OPEN_FLAG=--open"
if defined HOMEWORK_NO_OPEN set "OPEN_FLAG="

echo 正在排版，请稍等，大约半分钟……
echo.
".venv\Scripts\python.exe" -m homework_formatter %OPEN_FLAG%
set "CODE=%ERRORLEVEL%"
echo.
if "%CODE%"=="0" echo 做完了。
if not "%CODE%"=="0" echo 没有顺利完成，请看上面的提示。
if not "%CODE%"=="0" echo 如果上面是一串英文，请看《使用说明》里“常见问题”一节。
echo.
pause
exit /b %CODE%

:no_environment
echo 找不到运行环境：这个文件夹里没有 .venv 文件夹。
echo 请看《使用说明》里“换了电脑，或者移动了文件夹”一节。
echo.
pause
exit /b 1
