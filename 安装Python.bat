@echo off
chcp 65001
echo 正在检查Python环境...

set PYINSTALL=python-3.8.10-amd64.exe

:: 检查是否已安装Python
python --version >nul 2>&1
if not errorlevel 1 (
    echo 检测到Python已安装，跳过安装步骤
    goto :install_env
)

:: 1. 检查分发目录下是否有安装包
if exist "%PYINSTALL%" (
    echo 检测到本地已有Python安装包，跳过下载。
    set PYINSTALL_PATH=%CD%\%PYINSTALL%
    goto :install_python
)

:: 2. 检查temp目录下是否有安装包
if exist "temp\%PYINSTALL%" (
    echo 检测到temp目录已有Python安装包，跳过下载。
    set PYINSTALL_PATH=%CD%\temp\%PYINSTALL%
    goto :install_python
)

:: 3. 下载
mkdir temp 2>nul
cd temp
echo 未检测到本地安装包，开始下载...
powershell -Command "& {try {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe' -OutFile '%PYINSTALL%' -ErrorAction Stop} catch {Write-Host '下载失败，请检查网络连接'; exit 1}}"
if not exist "%PYINSTALL%" (
    echo 下载失败，请检查网络连接后重试
    cd ..
    pause
    exit /b 1
)
set PYINSTALL_PATH=%CD%\%PYINSTALL%
cd ..

:install_python
echo 正在安装Python...
"%PYINSTALL_PATH%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

:: 等待安装完成
echo 等待安装完成...
timeout /t 30 /nobreak

:: 验证安装
python --version >nul 2>&1
if errorlevel 1 (
    echo Python安装可能失败，请检查安装日志
    pause
    exit /b 1
)

:: 清理临时文件
if exist "temp\%PYINSTALL%" del "temp\%PYINSTALL%"
if exist temp rmdir /s /q temp

echo Python安装完成！

:install_env
echo 现在可以运行"安装环境.bat"来安装程序所需的其他依赖包
pause 