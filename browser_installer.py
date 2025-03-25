import sys
import subprocess
import logging
from PyQt5.QtWidgets import QProgressDialog, QMessageBox
from PyQt5.QtCore import Qt

logger = logging.getLogger(__name__)

class BrowserInstaller:
    def __init__(self, parent=None):
        self.parent = parent
        
    def install_playwright(self):
        """安装 Playwright 浏览器"""
        try:
            dialog = QProgressDialog("正在安装 Playwright 浏览器...", None, 0, 0, self.parent)
            dialog.setWindowModality(Qt.WindowModal)
            dialog.setWindowTitle("安装进度")
            dialog.setCancelButton(None)
            dialog.show()
            
            # 安装 Playwright
            subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
            
            dialog.close()
            QMessageBox.information(self.parent, "安装成功", "Playwright 浏览器安装完成！")
            return True
            
        except Exception as e:
            logger.error(f"安装 Playwright 失败: {str(e)}")
            QMessageBox.critical(self.parent, "安装失败", f"安装 Playwright 失败：{str(e)}")
            return False
            
    def install_undetected_chrome(self):
        """安装 Undetected Chrome"""
        try:
            dialog = QProgressDialog("正在安装 Undetected Chrome...", None, 0, 0, self.parent)
            dialog.setWindowModality(Qt.WindowModal)
            dialog.setWindowTitle("安装进度")
            dialog.setCancelButton(None)
            dialog.show()
            
            # 安装 undetected-chromedriver
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "undetected-chromedriver"])
            
            dialog.close()
            QMessageBox.information(self.parent, "安装成功", "Undetected Chrome 安装完成！")
            return True
            
        except Exception as e:
            logger.error(f"安装 Undetected Chrome 失败: {str(e)}")
            QMessageBox.critical(self.parent, "安装失败", f"安装 Undetected Chrome 失败：{str(e)}")
            return False
            
    def check_browser_installation(self, browser_type):
        """检查浏览器是否已安装"""
        try:
            if browser_type == "playwright":
                # 检查 Playwright 浏览器
                result = subprocess.run(
                    [sys.executable, "-m", "playwright", "install", "--help"],
                    capture_output=True,
                    text=True
                )
                return result.returncode == 0
                
            elif browser_type == "undetected":
                # 检查 undetected-chromedriver
                import undetected_chromedriver
                return True
                
            return True
            
        except Exception as e:
            logger.error(f"检查浏览器安装状态失败: {str(e)}")
            return False 