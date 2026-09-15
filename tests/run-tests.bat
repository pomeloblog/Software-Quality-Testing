@echo off
rem ============================================================
rem  run-tests.bat : one-click build + run all automated tests
rem
rem  Pre-conditions (see tests/README.md):
rem    1. Tomcat running with /Book deployed  -> http://localhost:8080/Book/
rem    2. MySQL running (schema `book`, user bookstore/123456)
rem
rem  Test scope (owner: member C / 丙):
rem    M3BookBrowseTest  = M3 book browsing/search  (BS-IT-041 ~ 046)
rem    M6ManagerTest     = M6 backend management    (BS-IT-101 ~ 106)
rem
rem  Note: a FAILED assertion here = an NG row in the Excel case
rem  list (defect evidence), not a script bug.
rem ============================================================
setlocal

rem UTF-8 codepage: JUnit failure messages contain Chinese (-Dfile.encoding=UTF-8)
chcp 65001 >nul

rem ---- stay inside tests\ ; every path below stays ASCII/relative ----
cd /d "%~dp0"

echo [1/3] Compiling tests ...
if exist classes rmdir /s /q classes
mkdir classes

rem only file NAMES go into the argfile (avoids GBK/UTF-8 argfile issue)
dir /b src\*.java > sources.txt

pushd src
javac -encoding UTF-8 -nowarn -cp "..\lib\*" -d "..\classes" "@..\sources.txt"
set "COMPILE_RC=%ERRORLEVEL%"
popd

if not "%COMPILE_RC%"=="0" (
    echo [ERROR] compilation failed.
    exit /b 1
)

echo [2/3] Running JUnit suite (12 cases) ...
echo.

java -Dfile.encoding=UTF-8 -cp "classes;lib\*" org.junit.runner.JUnitCore M3BookBrowseTest M6ManagerTest
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo [3/3] RESULT: ALL PASSED.
) else (
    echo [3/3] RESULT: FAILURES FOUND - each failure above is an NG entry in the Excel case list.
)
endlocal & exit /b %RC%
