@echo off
chcp 65001
echo 正在检查环境...

:: 检查Python是否已安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到Python，请先运行"安装Python.bat"
    pause
    exit /b 1
)

:: 检查是否已存在虚拟环境
if exist venv (
    echo 发现已存在的虚拟环境，是否重新创建？[Y/N]
    set /p choice=
    if /i "%choice%"=="Y" (
        echo 正在删除旧的虚拟环境...
        rmdir /s /q venv
    ) else (
        echo 使用现有虚拟环境...
        goto :activate_env
    )
)

:: 创建虚拟环境
echo 正在创建虚拟环境...
python -m venv venv
if errorlevel 1 (
    echo 创建虚拟环境失败，请检查Python安装
    pause
    exit /b 1
)

:activate_env
:: 激活虚拟环境
echo 正在激活虚拟环境...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo 激活虚拟环境失败
    pause
    exit /b 1
)

:: 检查pip是否可用
pip --version >nul 2>&1
if errorlevel 1 (
    echo 错误：pip不可用，请检查Python安装
    pause
    exit /b 1
)

:: 升级pip，使用清华镜像源
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo pip升级失败，请检查网络连接
    pause
    exit /b 1
)

:: 安装依赖，使用清华镜像源
echo 正在安装依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 安装依赖包失败，请检查网络连接或requirements.txt文件
    pause
    exit /b 1
)

echo 环境安装完成！
echo 请运行"启动工具.bat"来启动程序
pause 