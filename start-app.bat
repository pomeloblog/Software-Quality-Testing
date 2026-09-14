@echo off
setlocal
chcp 936 >nul

REM ==================================================================
REM  网上书店系统 - 一键启动（MySQL + Tomcat）
REM  作者：王尚清
REM
REM  会弹出两个新窗口分别运行 MySQL 与 Tomcat：
REM    - 请勿关闭那两个窗口，关掉即等于停止服务
REM    - 应用地址： http://localhost:8080/Book/
REM  若更换了运行目录，只需修改下面 RUNTIME 一行。
REM ==================================================================

set "RUNTIME=E:\网上书店项目依赖\runtime"
set "JAVA_HOME=%RUNTIME%\jdk8u504-b01"
set "CATALINA_HOME=%RUNTIME%\apache-tomcat-9.0.121"
set "CATALINA_BASE=%CATALINA_HOME%"

echo ==================================================================
echo   网上书店系统 - 启动 MySQL 与 Tomcat
echo ==================================================================
echo.

if not exist "%JAVA_HOME%\bin\java.exe" (
    echo [错误] 找不到 JDK：%JAVA_HOME%
    echo        请修改本脚本顶部的 RUNTIME 路径。
    goto :fail
)
if not exist "%RUNTIME%\mysql-5.7.44-winx64\bin\mysqld.exe" (
    echo [错误] 找不到 MySQL：%RUNTIME%\mysql-5.7.44-winx64\bin
    echo        请修改本脚本顶部的 RUNTIME 路径。
    goto :fail
)

netstat -ano | findstr ":3306" | findstr LISTENING >nul
if not errorlevel 1 (
    echo [提示] 3306 端口已被占用，MySQL 可能已经在运行，跳过启动。
) else (
    echo [1/3] 启动 MySQL ...
    start "MySQL - 网上书店" cmd /k %RUNTIME%\mysql-5.7.44-winx64\bin\mysqld.exe --datadir=%RUNTIME%\mysql-data --port=3306 --console
    call :waitPort 3306 30 "MySQL"
)

netstat -ano | findstr ":8080" | findstr LISTENING >nul
if not errorlevel 1 (
    echo [提示] 8080 端口已被占用，Tomcat 可能已经在运行，跳过启动。
) else (
    echo [2/3] 启动 Tomcat ...
    start "Tomcat - 网上书店" cmd /k %CATALINA_HOME%\bin\catalina.bat run
    call :waitPort 8080 60 "Tomcat"
)

echo [3/3] 检测应用是否就绪 ...
netstat -ano | findstr ":8080" | findstr LISTENING >nul
if errorlevel 1 (
    echo.
    echo [警告] 8080 端口尚未监听，Tomcat 可能还在启动或启动失败，
    echo        请查看 "Tomcat - 网上书店" 窗口中的错误信息。
) else (
    echo     应用已就绪，正在打开浏览器 ...
    start "" "http://localhost:8080/Book/"
)

echo.
echo ==================================================================
echo   应用地址： http://localhost:8080/Book/
echo   数据库：   localhost:3306  book  root/12345
echo   停止服务： 直接关闭那两个服务窗口
echo ==================================================================
echo.
pause
exit /b 0

REM ---------------------------------------------------------------
REM  :waitPort  端口号  最长等待秒数  服务名
REM  轮询端口直到监听为止，避免用固定 sleep 造成误判。
REM  注意：这里用 timeout.exe 的绝对路径，否则可能被 PATH 中
REM        其他同名程序（如 MSYS 的 timeout）抢先。
REM ---------------------------------------------------------------
REM  注意：这里用 ping 做延时而非 timeout.exe —— 后者在 stdin 被重定向时会直接报错，
REM        且必须用绝对路径，否则可能被 PATH 中其他同名程序（如 MSYS 的 timeout）抢先。
:waitPort
set "WP_PORT=%~1"
set "WP_MAX=%~2"
set "WP_NAME=%~3"
set /a WP_WAITED=0
:waitPortLoop
netstat -ano | findstr ":%WP_PORT%" | findstr LISTENING >nul
if not errorlevel 1 (
    echo     %WP_NAME% 已就绪（等待约 %WP_WAITED% 秒）
    goto :waitPortDone
)
set /a WP_WAITED+=2
if %WP_WAITED% geq %WP_MAX% (
    echo     [警告] %WP_NAME% 等待 %WP_MAX% 秒仍未就绪，请查看其窗口中的报错。
    goto :waitPortDone
)
%SystemRoot%\System32\ping.exe -n 3 127.0.0.1 >nul 2>&1
goto :waitPortLoop
:waitPortDone
exit /b 0

:fail
echo.
pause
exit /b 1
