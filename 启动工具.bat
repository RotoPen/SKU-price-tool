@echo off
chcp 65001
echo 正在启动 SKU活动价自动匹配与审核工具...

:: 如果存在虚拟环境，则激活它
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

:: 启动程序
streamlit run "sku_price_checker.py" 