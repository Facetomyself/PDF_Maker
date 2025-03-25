import os
import pandas as pd
from datetime import datetime
import uuid
import base64
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from browser_manager import BrowserManager
import re

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class PDFMaker:
    def __init__(self, config):
        self.config = config
        self.chrome_path = config.get('paths', 'chrome_path')
        self.template_path = config.get('paths', 'template_path')
        self.output_dir = config.get('paths', 'output_dir')
        self.browser_manager = BrowserManager(config)
        
        # Excel 字段到 HTML 占位符的映射关系
        self.field_mapping = {
            # Excel字段名: HTML占位符名
            'order id': '{{order_id}}',
            'total invoice value': '{{total_value}}',
            'update time': '{{update_time}}',
            'telephone': '{{telephone}}',
            'country': '{{country}}',
            'address1': '{{address1}}',
            'address2': '{{address2}}'
        }
        
        # 确保输出目录存在
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def read_excel(self, excel_file):
        """读取 Excel 文件"""
        df = pd.read_excel(excel_file)
        logger.info(f"成功读取 Excel 文件，列名: {list(df.columns)}")
        return df

    def format_value(self, value):
        """格式化值，处理不同的数据类型"""
        if pd.isna(value):  # 处理空值
            return ""
        elif isinstance(value, (int, float)):  # 处理数字
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
        return str(value)

    def render_template(self, row_data, field_mapping):
        """渲染 HTML 模板"""
        try:
            with open(self.template_path, 'r', encoding='utf-8') as f:
                template_content = f.read()
            
            # 查找所有占位符
            placeholders = re.findall(r'\{\{([^}]+)\}\}', template_content)
            logger.debug(f"模板中的占位符: {placeholders}")
            
            # 创建替换映射
            replacements = {}
            for placeholder in placeholders:
                # 检查是否有对应的 Excel 字段
                excel_field = None
                for field, mapped_placeholder in field_mapping.items():
                    if mapped_placeholder == placeholder:
                        excel_field = field
                        break
                
                if excel_field and excel_field in row_data:
                    value = self.format_value(row_data[excel_field])
                    replacements[f"{{{{{placeholder}}}}}"] = value
                    logger.debug(f"替换占位符 {placeholder} 为值: {value}")
                else:
                    replacements[f"{{{{{placeholder}}}}}"] = ""
                    logger.debug(f"未找到占位符 {placeholder} 的映射，设置为空")
            
            # 应用所有替换
            for old, new in replacements.items():
                template_content = template_content.replace(old, new)
            
            return template_content
            
        except Exception as e:
            logger.error(f"渲染模板时发生错误: {str(e)}")
            raise

    def get_chrome_driver(self):
        """配置并返回 Chrome driver"""
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.binary_location = self.chrome_path
        
        service = Service()
        driver = webdriver.Chrome(service=service, options=options)
        return driver

    def print_to_pdf(self, driver, options=None):
        """使用 CDP 命令直接打印为 PDF"""
        if options is None:
            options = {}

        default_options = {
            'paperWidth': self.config.get('pdf_settings', 'paper_width'),
            'paperHeight': self.config.get('pdf_settings', 'paper_height'),
            'marginTop': self.config.get('pdf_settings', 'margin_top'),
            'marginBottom': self.config.get('pdf_settings', 'margin_bottom'),
            'marginLeft': self.config.get('pdf_settings', 'margin_left'),
            'marginRight': self.config.get('pdf_settings', 'margin_right'),
            'printBackground': True,
            'preferCSSPageSize': True,
            'scale': self.config.get('pdf_settings', 'scale')
        }
        default_options.update(options)
        
        result = driver.execute_cdp_cmd('Page.printToPDF', default_options)
        return base64.b64decode(result['data'])

    def generate_pdf(self, html_content, row_data):
        """生成 PDF 文件"""
        try:
            # 生成唯一的输出文件名（使用订单号和UUID）
            try:
                order_id = self.format_value(row_data['平台订单号'])
            except KeyError:
                order_id = datetime.now().strftime('%Y%m%d%H%M%S')
                
            unique_id = uuid.uuid4().hex[:8]
            output_file = os.path.join(self.output_dir, f"order_{order_id}_{unique_id}.pdf")
            
            # 使用浏览器管理器生成 PDF
            self.browser_manager.print_to_pdf(html_content, output_file)
            logger.info(f"成功生成 PDF 文件: {output_file}")
            return output_file
            
        except Exception as e:
            logger.error(f"生成 PDF 时发生错误: {str(e)}")
            return None

    def process(self, excel_file):
        """处理整个流程"""
        try:
            logger.info(f"开始处理 Excel 文件: {excel_file}")
            
            # 读取 Excel 文件
            df = self.read_excel(excel_file)
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    logger.info(f"正在处理第 {index + 1} 行数据...")
                    
                    # 渲染模板
                    html_content = self.render_template(row, self.field_mapping)
                    
                    # 生成 PDF
                    output_file = self.generate_pdf(html_content, row)
                    
                    if output_file:
                        logger.info(f"第 {index + 1} 行 PDF 生成成功: {output_file}")
                    else:
                        logger.error(f"第 {index + 1} 行 PDF 生成失败")
                        
                except Exception as e:
                    logger.error(f"处理第 {index + 1} 行时发生错误: {str(e)}")
                    continue
                
        except Exception as e:
            logger.error(f"处理过程中发生错误: {str(e)}")

if __name__ == "__main__":
    pdf_maker = PDFMaker()
    pdf_maker.process("PayOrder_1742629289639.xlsx")
