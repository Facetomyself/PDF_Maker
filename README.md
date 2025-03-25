# PDF生成器 (PDF Maker)

一个基于PyQt5的PDF生成工具，支持从Excel数据生成自定义模板的PDF文件。

## 功能特点

- 📊 支持Excel数据源导入
- 🎨 自定义HTML模板
- 🔄 字段映射系统
- 🌐 多浏览器引擎支持
  - 本地浏览器
  - Undetected Chrome
  - Playwright
- ⏯️ 支持暂停/继续/停止生成过程
- 📝 详细的日志记录
- 🎯 简单直观的用户界面

## 系统要求

- Python 3.7+
- Windows 7/10/11

## 依赖项

主要依赖包：
```python
PyQt5
pandas
openpyxl
selenium
playwright
undetected-chromedriver
pdfkit
loguru
```

## 安装说明

1. 克隆仓库
```bash
git clone https://github.com/Facetomyself/PDF_Maker.git
cd PDF_Maker
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 运行程序
```bash
python main.py
```

## 使用说明

1. 选择Excel文件：点击"选择Excel文件"按钮，选择包含数据的Excel文件
2. 选择HTML模板：点击"选择模板文件"按钮，选择HTML模板文件
3. 配置字段映射：将Excel列名与HTML模板中的占位符进行映射
4. 选择浏览器引擎：根据需要选择合适的浏览器引擎
5. 点击"生成PDF"开始生成过程

## 目录结构 

PDF_Maker/
├── main.py # 主程序
├── config_manager.py # 配置管理
├── PDF_Maker.py # PDF生成核心
├── browser_installer.py # 浏览器安装器
├── logger_manager.py # 日志管理
├── utils.py # 工具函数
├── build.py # 打包脚本
├── config.xml # 配置文件
├── logs/ # 日志目录
└── finished/ # 输出目录