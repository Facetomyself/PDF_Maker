import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
from playwright.sync_api import sync_playwright
import tempfile
import shutil

class BrowserManager:
    def __init__(self, config):
        self.config = config
        self.browser_type = config.get('browser', 'type', default='local')
        self.chrome_path = config.get('paths', 'chrome_path')
        self.temp_dir = None
        
    def create_temp_dir(self):
        """创建临时目录用于存储浏览器文件"""
        if not self.temp_dir:
            self.temp_dir = tempfile.mkdtemp(prefix='chrome_')
        return self.temp_dir
        
    def cleanup(self):
        """清理临时文件"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"清理临时文件失败: {str(e)}")
                
    def get_selenium_driver(self):
        """获取 Selenium WebDriver"""
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.binary_location = self.chrome_path
        
        service = Service()
        driver = webdriver.Chrome(service=service, options=options)
        return driver
        
    def get_undetected_driver(self):
        """获取 Undetected ChromeDriver"""
        options = uc.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        driver = uc.Chrome(
            options=options,
            browser_executable_path=self.chrome_path,
            user_data_dir=self.create_temp_dir()
        )
        return driver
        
    def get_playwright_browser(self):
        """获取 Playwright 浏览器"""
        playwright = sync_playwright().start()
        browser = playwright.chromium.launch(
            headless=True,
            executable_path=self.chrome_path
        )
        return browser, playwright
        
    def get_browser(self):
        """根据配置获取浏览器实例"""
        if self.browser_type == 'local':
            return self.get_selenium_driver()
        elif self.browser_type == 'undetected':
            return self.get_undetected_driver()
        elif self.browser_type == 'playwright':
            return self.get_playwright_browser()
        else:
            raise ValueError(f"不支持的浏览器类型: {self.browser_type}")
            
    def print_to_pdf(self, html_content, output_path):
        """使用选定的浏览器打印 PDF"""
        try:
            if self.browser_type == 'playwright':
                browser, playwright = self.get_playwright_browser()
                try:
                    page = browser.new_page()
                    page.set_content(html_content)
                    page.pdf(path=output_path, format='A4')
                finally:
                    browser.close()
                    playwright.stop()
            else:
                # 创建临时 HTML 文件
                temp_html = os.path.join(self.create_temp_dir(), 'temp.html')
                with open(temp_html, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                    
                # 获取浏览器实例
                if self.browser_type == 'local':
                    driver = self.get_selenium_driver()
                else:
                    driver = self.get_undetected_driver()
                    
                try:
                    # 加载 HTML 文件
                    driver.get(f'file:///{os.path.abspath(temp_html)}')
                    
                    # 打印为 PDF
                    pdf_data = driver.execute_cdp_cmd('Page.printToPDF', {
                        'paperWidth': self.config.get('pdf_settings', 'paper_width'),
                        'paperHeight': self.config.get('pdf_settings', 'paper_height'),
                        'marginTop': self.config.get('pdf_settings', 'margin_top'),
                        'marginBottom': self.config.get('pdf_settings', 'margin_bottom'),
                        'marginLeft': self.config.get('pdf_settings', 'margin_left'),
                        'marginRight': self.config.get('pdf_settings', 'margin_right'),
                        'printBackground': True,
                        'preferCSSPageSize': True,
                        'scale': self.config.get('pdf_settings', 'scale')
                    })
                    
                    # 保存 PDF 文件
                    with open(output_path, 'wb') as f:
                        f.write(pdf_data['data'])
                        
                finally:
                    driver.quit()
                    
        finally:
            self.cleanup() 