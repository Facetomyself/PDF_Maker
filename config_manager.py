import xml.etree.ElementTree as ET
import os

class ConfigManager:
    def __init__(self):
        self.config_file = 'config.xml'
        self.tree = None
        self.root = None
        self.load_config()
        
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            self.tree = ET.parse(self.config_file)
            self.root = self.tree.getroot()
        else:
            self.create_default_config()
            
    def create_default_config(self):
        """创建默认配置文件"""
        self.root = ET.Element('config')
        
        # 路径配置
        paths = ET.SubElement(self.root, 'paths')
        ET.SubElement(paths, 'chrome_path').text = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        ET.SubElement(paths, 'template_path').text = "tr91vewuqjsieb6ot1zd.html"
        ET.SubElement(paths, 'output_dir').text = "finished"
        
        # PDF 设置
        pdf_settings = ET.SubElement(self.root, 'pdf_settings')
        ET.SubElement(pdf_settings, 'paper_width').text = "8.27"
        ET.SubElement(pdf_settings, 'paper_height').text = "11.69"
        ET.SubElement(pdf_settings, 'margin_top').text = "0"
        ET.SubElement(pdf_settings, 'margin_bottom').text = "0"
        ET.SubElement(pdf_settings, 'margin_left').text = "0"
        ET.SubElement(pdf_settings, 'margin_right').text = "0"
        ET.SubElement(pdf_settings, 'scale').text = "1.0"
        
        # 浏览器设置
        browser = ET.SubElement(self.root, 'browser')
        ET.SubElement(browser, 'type').text = "local"  # local, undetected, playwright
        
        self.tree = ET.ElementTree(self.root)
        self.save_config()
        
    def save_config(self):
        """保存配置到文件"""
        self.tree.write(self.config_file, encoding='utf-8', xml_declaration=True)
        
    def get(self, section, key, default=None):
        """获取配置值"""
        element = self.root.find(f'.//{section}/{key}')
        return element.text if element is not None else default
        
    def set(self, section, key, value):
        """设置配置值"""
        section_element = self.root.find(f'.//{section}')
        if section_element is None:
            section_element = ET.SubElement(self.root, section)
            
        key_element = section_element.find(key)
        if key_element is None:
            key_element = ET.SubElement(section_element, key)
            
        key_element.text = str(value)
        self.save_config() 