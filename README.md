# SKU活动价自动匹配与审核工具

一个基于Streamlit的自动化工具，用于匹配和审核SKU活动价格。

## 功能特点

- 自动匹配SKU活动价格
- 支持Parent SKU匹配
- 人工审核功能
- 导出最终价格表
- 可视化高亮显示

## 安装与使用

1. 安装依赖：
   ```
   pip install -r requirements.txt
   ```

2. 运行程序：
   ```
   streamlit run main.py
   ```

## 使用说明

1. 上传SKU表、工具价格表、活动价格提交表
2. 系统自动匹配价格
3. 人工确认或修改价格
4. 导出最终表格

## 依赖项

- streamlit
- pandas
- numpy
- openpyxl
- st-aggrid 