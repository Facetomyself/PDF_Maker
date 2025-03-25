import PyInstaller.__main__
import os
import shutil
import sys
from datetime import datetime
import importlib

def create_directories():
    """创建必要的目录"""
    directories = ['logs', 'finished']
    for dir_name in directories:
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
            print(f"创建目录：{dir_name}")

def clean_build_files():
    """清理构建文件"""
    # 清理目录
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"清理目录：{dir_name}")
    
    # 清理 spec 文件
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            os.remove(file)
            print(f"清理文件：{file}")

def copy_resources(dist_dir):
    """复制资源文件到目标目录"""
    resources = [
        ('config.xml', '配置文件'),
        ('tr91vewuqjsieb6ot1zd.html', '模板html文件'),
        ('PayOrder_1742629289639.xlsx', '模板Excel文件'),
        ('logs', '日志目录'),
        ('finished', '输出目录')
    ]
    
    for resource, desc in resources:
        if os.path.exists(resource):
            if os.path.isdir(resource):
                shutil.copytree(resource, os.path.join(dist_dir, resource))
            else:
                shutil.copy2(resource, dist_dir)
            print(f"复制{desc}：{resource}")
        else:
            print(f"警告：{desc} {resource} 不存在")

def build():
    """构建可执行文件"""
    print("开始构建 PDF 生成器...")
    
    # 清理之前的构建文件
    clean_build_files()
    
    # 创建必要的目录
    create_directories()
    
    # PyInstaller 参数
    params = [
        'main.py',  # 主程序文件
        '--name=PDF生成器',  # 生成的exe名称
        '--windowed',  # 使用 GUI 模式
        '--icon=icon.ico',  # 程序图标
        '--add-data=config.xml;.',  # 添加配置文件
        '--add-data=tr91vewuqjsieb6ot1zd.html;.',  # 添加模板文件
        '--add-data=logs;logs',  # 添加日志目录
        '--add-data=finished;finished',  # 添加输出目录
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
        '--noconsole',  # 不显示控制台窗口
        # 核心依赖
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=PyQt5',
        '--hidden-import=loguru',
        '--hidden-import=utils',  # 添加工具模块
        # 浏览器相关
        '--hidden-import=selenium',
        '--hidden-import=playwright',
        '--hidden-import=undetected_chromedriver',
        '--hidden-import=webdriver_manager',
        # PDF相关
        '--hidden-import=pdfkit',
        '--hidden-import=wkhtmltopdf',
        # 添加必要的包数据
        f'--add-data={os.path.dirname(importlib.import_module("pandas").__file__)};pandas',
        f'--add-data={os.path.dirname(importlib.import_module("openpyxl").__file__)};openpyxl',
        f'--add-data={os.path.dirname(importlib.import_module("PyQt5").__file__)};PyQt5',
        f'--add-data={os.path.dirname(importlib.import_module("loguru").__file__)};loguru',
        f'--add-data={os.path.dirname(importlib.import_module("selenium").__file__)};selenium',
        f'--add-data={os.path.dirname(importlib.import_module("playwright").__file__)};playwright',
        f'--add-data={os.path.dirname(importlib.import_module("undetected_chromedriver").__file__)};undetected_chromedriver',
        '--collect-all=webdriver_manager',
        '--collect-all=pdfkit',
    ]

    # 添加调试信息收集
    params.extend([
        '--debug=all',  # 启用所有调试信息
        '--log-level=DEBUG',  # 设置日志级别
    ])
    
    # 执行打包
    print("正在打包...")
    PyInstaller.__main__.run(params)
    
    # 复制额外文件到dist目录
    dist_dir = os.path.join('dist', 'PDF生成器')
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)
        
    # 复制资源文件
    copy_resources(dist_dir)
    
    # 清理 build 文件夹
    if os.path.exists('build'):
        shutil.rmtree('build')
        print("清理构建临时文件：build")
    
    print("\n打包完成！")
    print(f"可执行文件位于：{os.path.abspath(dist_dir)}")
    print(f"构建时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("\n注意：")
    print("1. 最终程序在 dist/PDF生成器 目录下")
    print("2. 可以直接分发 dist/PDF生成器 目录")
    print("3. build 文件夹已自动清理")
    print("4. 请确保 finished 和 logs 目录存在且有写入权限")

if __name__ == '__main__':
    try:
        build()
    except Exception as e:
        print(f"构建过程中发生错误：{str(e)}")
        sys.exit(1) 