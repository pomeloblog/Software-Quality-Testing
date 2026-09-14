@echo off
setlocal
chcp 936 >nul

REM ==================================================================
REM  网上书店系统 - 自动化测试 一键运行脚本
REM  被测模块：M1 用户认证 / M4 购物车      作者：王尚清
REM
REM  运行前提：MySQL 与 Tomcat 已启动（本脚本只跑测试，不负责启动应用）
REM  用法：双击本文件，或在项目根目录执行  run-tests.bat
REM ==================================================================

REM ---------- 运行环境（如与你的机器不同，改这两行即可）----------
set "JAVA_HOME=E:\网上书店项目依赖\runtime\jdk8u504-b01"
set "TOMCAT_HOME=E:\网上书店项目依赖\runtime\apache-tomcat-9.0.121"

cd /d "%~dp0"
set "BUILD=%~dp0build"

echo.
echo ==================================================================
echo   网上书店系统 自动化测试
echo ==================================================================
echo.

echo [1/5] 检查运行环境
if not exist "%JAVA_HOME%\bin\javac.exe" (
    echo     [错误] 找不到 JDK：%JAVA_HOME%
    echo     请修改本脚本顶部的 JAVA_HOME 路径。
    goto :fail
)
if not exist "%TOMCAT_HOME%\lib\servlet-api.jar" (
    echo     [错误] 找不到 servlet-api.jar：%TOMCAT_HOME%\lib
    echo     请修改本脚本顶部的 TOMCAT_HOME 路径。
    goto :fail
)
echo     JAVA_HOME   = %JAVA_HOME%
echo     TOMCAT_HOME = %TOMCAT_HOME%
"%JAVA_HOME%\bin\java.exe" -version 2>&1 | findstr /i "version"

echo.
echo [2/5] 准备依赖
if not exist "%BUILD%\classes" mkdir "%BUILD%\classes"
if not exist "%BUILD%\lib" mkdir "%BUILD%\lib"
copy /y "%TOMCAT_HOME%\lib\servlet-api.jar" "%BUILD%\lib\" >nul
copy /y "%TOMCAT_HOME%\lib\jsp-api.jar" "%BUILD%\lib\" >nul
copy /y "src\jdbc.properties" "%BUILD%\classes\" >nul
if errorlevel 1 goto :fail

echo.
echo [3/5] 生成源文件清单
dir /s /b src\com\*.java > "%BUILD%\sources.txt"

echo.
echo [4/5] 编译被测代码与测试代码
"%JAVA_HOME%\bin\javac.exe" -encoding UTF-8 -source 1.8 -target 1.8 -nowarn -cp "%BUILD%\lib\*;web\WEB-INF\lib\*" -d "%BUILD%\classes" "@%BUILD%\sources.txt"
if errorlevel 1 goto :fail
echo     编译完成

echo.
echo [5/5] 执行测试
echo.
"%JAVA_HOME%\bin\java.exe" -cp "%BUILD%\classes;%BUILD%\lib\*;web\WEB-INF\lib\*" com.yj.test.TestRunner

echo.
echo 提示：
echo   - 用例编号与《测试用例清单》一一对应，逐条结果可直接复制到测试报告的「执行日志」。
echo   - 未通过的用例代表仍有未修复的缺陷；全部通过则说明缺陷已修复且未回归。
echo.
pause
exit /b 0

:fail
echo.
echo ==================================================================
echo   构建失败，请根据上方错误信息排查。
echo ==================================================================
echo.
pause
exit /b 1
