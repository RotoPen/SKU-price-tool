# SKU活动价自动匹配与审核工具

一个基于Streamlit的自动化工具，支持SKU活动价格的批量匹配、人工审核与导出，适用于电商运营、价格管理等场景。

## 在线体验

> [点击访问云端演示（Streamlit Cloud）](https://sku-price-tool-hxopmlanacmposbxv9tb2n.streamlit.app/)

无需本地安装，直接网页操作。

## 功能特点

- **自动匹配SKU活动价格**：优先匹配工具价格表，支持Parent SKU回退，最终使用推荐价格。
- **人工审核与修改**：可对推荐价格进行人工确认或手动调整。
- **高亮可视化**：不同价格来源高亮显示，异常/缺失价格红色警示。
- **一键导出**：支持导出带有价格标记的最终活动价格表（Excel）。
- **灵活配置**：可自定义价格浮动范围，支持备注行跳过。

## 安装与本地运行

1. 克隆仓库：
   ```bash
   git clone https://github.com/你的用户名/sku-price-tool.git
   cd sku-price-tool
   ```
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 启动应用：
   ```bash
   streamlit run sku_price_checker.py
   ```

---

## 使用批处理（BAT）文件本地运行教程（推荐给Windows用户）

本项目已为Windows用户准备了全套一键批处理脚本，无需手动配置环境，适合零基础用户。

### 步骤一：自动安装Python（如本机未安装）

1. 双击 `安装Python.bat`
   - 自动检测本机是否已安装Python。
   - 如未安装，会自动下载并静默安装Python 3.8（需联网）。
   - 安装完成后，命令行会提示继续下一步。

### 步骤二：自动创建虚拟环境并安装依赖

2. 双击 `安装环境.bat`
   - 自动创建Python虚拟环境（venv）。
   - 自动激活虚拟环境，并用清华镜像源安装所有依赖包。
   - 安装完成后，命令行会提示"请运行'启动工具.bat'来启动程序"。

### 步骤三：一键启动工具

3. 双击 `启动工具.bat`
   - 自动激活虚拟环境并启动工具，无需手动输入命令。
   - 程序会自动在浏览器中打开Streamlit网页界面。

> **注意事项：**
> - 所有脚本均可直接在资源管理器中双击运行，无需命令行基础。
> - 如遇到"权限不足"或"无法运行脚本"问题，请右键以管理员身份运行，或检查杀毒软件拦截。
> - 推荐将项目文件夹路径中不要包含中文或空格，以避免部分环境下的兼容性问题。

## 依赖环境

- streamlit >= 1.24.0
- pandas >= 1.5.0
- numpy >= 1.21.0
- openpyxl >= 3.0.0
- streamlit-aggrid == 0.3.4.post3
- 其它见 requirements.txt

## 使用说明

1. 上传SKU表、工具价格表、活动价格提交表（支持Excel/CSV）
2. 系统自动匹配并填写活动价格
3. 对推荐价格可人工确认或修改
4. 导出最终活动价格表（Excel）

## 云端部署说明

- 本项目已适配 [Streamlit Cloud](https://streamlit.io/cloud)
- 只需将代码推送到GitHub公开仓库，云端即可一键部署
- requirements.txt已锁定兼容依赖版本
