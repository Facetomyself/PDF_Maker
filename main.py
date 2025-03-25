import sys
import re
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QListWidget, 
                           QFileDialog, QMessageBox, QProgressBar, QGroupBox, QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import os
from config_manager import ConfigManager
from PDF_Maker import PDFMaker
from browser_installer import BrowserInstaller
from logger_manager import LoggerManager

class PDFGeneratorThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    status_changed = pyqtSignal(str)  # 新增状态信号
    
    def __init__(self, excel_file, field_mapping, config):
        super().__init__()
        self.excel_file = excel_file
        self.field_mapping = field_mapping
        self.config = config
        self.is_paused = False
        self.is_stopped = False
        
    def pause(self):
        """暂停处理"""
        self.is_paused = True
        self.status_changed.emit("已暂停")
        
    def resume(self):
        """继续处理"""
        self.is_paused = False
        self.status_changed.emit("正在处理")
        
    def stop(self):
        """停止处理"""
        self.is_stopped = True
        self.status_changed.emit("已停止")
        
    def run(self):
        logger = LoggerManager().get_logger()
        try:
            # 读取 Excel 文件
            logger.info(f"开始读取 Excel 文件：{self.excel_file}")
            df = pd.read_excel(self.excel_file)
            total_rows = len(df)
            logger.info(f"Excel 文件读取成功，共 {total_rows} 行数据")
            
            # 创建 PDF 生成器
            pdf_maker = PDFMaker(self.config)
            
            # 处理每一行数据
            for index, row in df.iterrows():
                if self.is_stopped:
                    logger.warning("用户手动停止生成过程")
                    self.status_changed.emit("已停止")
                    break
                    
                while self.is_paused:
                    self.msleep(100)  # 暂停时等待
                    if self.is_stopped:
                        break
                
                if self.is_stopped:
                    break
                
                # 渲染模板
                logger.debug(f"正在处理第 {index + 1}/{total_rows} 行数据")
                html_content = pdf_maker.render_template(row, self.field_mapping)
                
                # 生成 PDF
                pdf_maker.generate_pdf(html_content, row)
                
                # 更新进度
                progress = int((index + 1) / total_rows * 100)
                self.progress.emit(progress)
            
            if not self.is_stopped:
                logger.info("PDF 生成完成")
                self.finished.emit()
            
        except Exception as e:
            logger.error(f"生成 PDF 时发生错误：{str(e)}")
            self.error.emit(str(e))

class MappingPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 预览标题
        title = QLabel("映射预览")
        title.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(title)
        
        # 预览内容
        self.preview_text = QLabel()
        self.preview_text.setWordWrap(True)
        self.preview_text.setStyleSheet("""
            QLabel {
                background-color: #f5f5f5;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.preview_text)
        
    def update_preview(self, field_mapping):
        if not field_mapping:
            self.preview_text.setText("暂无映射关系")
            return
            
        preview_html = "<div style='font-family: Arial;'>"
        for excel_field, placeholder in field_mapping.items():
            preview_html += f"<p><b>{excel_field}</b> → <code>{placeholder}</code></p>"
        preview_html += "</div>"
        self.preview_text.setText(preview_html)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logger = LoggerManager().get_logger()
        self.config = ConfigManager()
        self.field_mapping = {}
        self.browser_installer = BrowserInstaller(self)
        self.setup_ui()
        self.logger.info("PDF 生成器启动")
        
    def setup_ui(self):
        self.setWindowTitle("PDF 生成器")
        self.setMinimumSize(800, 600)
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QVBoxLayout(main_widget)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # Excel 文件选择
        excel_layout = QHBoxLayout()
        self.excel_label = QLabel("Excel 文件：未选择")
        excel_btn = QPushButton("选择 Excel 文件")
        excel_btn.clicked.connect(self.select_excel)
        excel_layout.addWidget(self.excel_label)
        excel_layout.addWidget(excel_btn)
        file_layout.addLayout(excel_layout)
        
        # 模板文件选择
        template_layout = QHBoxLayout()
        self.template_label = QLabel("HTML 模板：未选择")
        template_btn = QPushButton("选择模板文件")
        template_btn.clicked.connect(self.select_template)
        template_layout.addWidget(self.template_label)
        template_layout.addWidget(template_btn)
        file_layout.addLayout(template_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 映射区域
        mapping_group = QGroupBox("字段映射")
        mapping_layout = QHBoxLayout()
        
        # Excel 字段列表
        excel_list_group = QGroupBox("Excel 字段")
        excel_list_layout = QVBoxLayout()
        self.excel_list = QListWidget()
        self.excel_list.setSelectionMode(QListWidget.SingleSelection)
        excel_list_layout.addWidget(self.excel_list)
        excel_list_group.setLayout(excel_list_layout)
        
        # 映射按钮
        mapping_buttons_layout = QVBoxLayout()
        self.map_btn = QPushButton("→")
        self.map_btn.clicked.connect(self.create_mapping)
        self.unmap_btn = QPushButton("←")
        self.unmap_btn.clicked.connect(self.remove_mapping)
        mapping_buttons_layout.addWidget(self.map_btn)
        mapping_buttons_layout.addWidget(self.unmap_btn)
        mapping_buttons_layout.addStretch()
        
        # HTML 占位符列表
        html_list_group = QGroupBox("HTML 占位符")
        html_list_layout = QVBoxLayout()
        self.html_list = QListWidget()
        self.html_list.setSelectionMode(QListWidget.SingleSelection)
        html_list_layout.addWidget(self.html_list)
        html_list_group.setLayout(html_list_layout)
        
        # 添加列表到映射布局
        mapping_layout.addWidget(excel_list_group)
        mapping_layout.addLayout(mapping_buttons_layout)
        mapping_layout.addWidget(html_list_group)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 映射预览
        preview_group = QGroupBox("映射预览")
        preview_layout = QVBoxLayout()
        self.preview_widget = MappingPreviewWidget()
        preview_layout.addWidget(self.preview_widget)
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # 按钮区域
        button_group = QGroupBox("控制面板")
        button_layout = QHBoxLayout()
        
        # 生成按钮
        self.generate_btn = QPushButton("生成 PDF")
        self.generate_btn.clicked.connect(self.generate_pdfs)
        button_layout.addWidget(self.generate_btn)
        
        # 暂停按钮
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.setEnabled(False)
        button_layout.addWidget(self.pause_btn)
        
        # 停止按钮
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_generation)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.stop_btn)
        
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        layout.addWidget(self.status_label)
        
        # 浏览器选择区域
        browser_group = QGroupBox("浏览器设置")
        browser_layout = QHBoxLayout()
        
        browser_label = QLabel("浏览器类型：")
        self.browser_combo = QComboBox()
        self.browser_combo.addItems(["本地浏览器", "Undetected Chrome", "Playwright"])
        self.browser_combo.currentIndexChanged.connect(self.change_browser)
        browser_layout.addWidget(browser_label)
        browser_layout.addWidget(self.browser_combo)
        browser_layout.addStretch()
        
        browser_group.setLayout(browser_layout)
        layout.addWidget(browser_group)
        
        # 设置样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #cccccc;
                border-radius: 4px;
                margin-top: 1em;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
            QListWidget {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
            }
        """)
        
    def select_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.logger.info(f"选择 Excel 文件：{file_name}")
            self.excel_label.setText(f"Excel 文件：{os.path.basename(file_name)}")
            self.excel_file = file_name
            self.update_excel_fields()
            
    def select_template(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择 HTML 模板", "", "HTML Files (*.html)")
        if file_name:
            self.logger.info(f"选择模板文件：{file_name}")
            self.template_label.setText(f"HTML 模板：{os.path.basename(file_name)}")
            self.config.set('paths', 'template_path', file_name)
            self.update_html_placeholders()
            
    def update_excel_fields(self):
        try:
            df = pd.read_excel(self.excel_file)
            self.excel_list.clear()
            for column in df.columns:
                self.excel_list.addItem(column)
            self.logger.info(f"更新 Excel 字段列表，共 {len(df.columns)} 个字段")
        except Exception as e:
            self.logger.error(f"读取 Excel 文件失败：{str(e)}")
            QMessageBox.critical(self, "错误", f"读取 Excel 文件失败：{str(e)}")
            
    def update_html_placeholders(self):
        try:
            with open(self.config.get('paths', 'template_path'), 'r', encoding='utf-8') as f:
                content = f.read()
                
            # 查找所有 {{xxx}} 格式的占位符
            placeholders = re.findall(r'\{\{([^}]+)\}\}', content)
            
            self.html_list.clear()
            for placeholder in placeholders:
                self.html_list.addItem(placeholder)
            
            self.logger.info(f"更新 HTML 占位符列表，共 {len(placeholders)} 个占位符")
                
        except Exception as e:
            self.logger.error(f"读取模板文件失败：{str(e)}")
            QMessageBox.critical(self, "错误", f"读取模板文件失败：{str(e)}")
            
    def create_mapping(self):
        excel_item = self.excel_list.currentItem()
        html_item = self.html_list.currentItem()
        
        if excel_item and html_item:
            excel_field = excel_item.text()
            html_placeholder = html_item.text()
            
            self.field_mapping[excel_field] = html_placeholder
            self.logger.info(f"创建映射：{excel_field} -> {html_placeholder}")
            
            # 更新预览
            self.preview_widget.update_preview(self.field_mapping)
            
            # 禁用已映射的项
            excel_item.setFlags(excel_item.flags() & ~Qt.ItemIsEnabled)
            html_item.setFlags(html_item.flags() & ~Qt.ItemIsEnabled)
            
    def remove_mapping(self):
        excel_item = self.excel_list.currentItem()
        html_item = self.html_list.currentItem()
        
        if excel_item and html_item:
            excel_field = excel_item.text()
            html_placeholder = html_item.text()
            
            if excel_field in self.field_mapping:
                del self.field_mapping[excel_field]
                self.logger.info(f"移除映射：{excel_field} -> {html_placeholder}")
                
                # 更新预览
                self.preview_widget.update_preview(self.field_mapping)
                
                # 重新启用项
                excel_item.setFlags(excel_item.flags() | Qt.ItemIsEnabled)
                html_item.setFlags(html_item.flags() | Qt.ItemIsEnabled)
                
    def generate_pdfs(self):
        if not hasattr(self, 'excel_file'):
            self.logger.warning("未选择 Excel 文件")
            QMessageBox.warning(self, "警告", "请先选择 Excel 文件")
            return
            
        if not self.field_mapping:
            self.logger.warning("未设置字段映射")
            QMessageBox.warning(self, "警告", "请先设置字段映射")
            return
            
        self.logger.info("开始生成 PDF")
        # 更新按钮状态
        self.generate_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setText("暂停")
        self.progress_bar.setValue(0)
        self.status_label.setText("正在生成 PDF...")
        
        # 创建并启动生成线程
        self.generator_thread = PDFGeneratorThread(
            self.excel_file, self.field_mapping, self.config)
        self.generator_thread.progress.connect(self.update_progress)
        self.generator_thread.finished.connect(self.generation_finished)
        self.generator_thread.error.connect(self.generation_error)
        self.generator_thread.status_changed.connect(self.update_status)
        self.generator_thread.start()
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def generation_finished(self):
        self.generate_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("PDF 生成完成！")
        QMessageBox.information(self, "完成", "所有 PDF 文件已生成完成！")
        
    def generation_error(self, error_msg):
        self.generate_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("生成失败")
        QMessageBox.critical(self, "错误", f"生成 PDF 时发生错误：{error_msg}")

    def change_browser(self, index):
        """切换浏览器类型"""
        browser_types = {
            0: "local",
            1: "undetected",
            2: "playwright"
        }
        
        browser_type = browser_types[index]
        self.logger.info(f"切换浏览器类型：{browser_type}")
        
        # 检查浏览器是否已安装
        if browser_type != "local" and not self.browser_installer.check_browser_installation(browser_type):
            reply = QMessageBox.question(
                self,
                "安装提示",
                f"需要安装 {browser_type.title()} 浏览器，是否现在安装？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                if browser_type == "playwright":
                    if not self.browser_installer.install_playwright():
                        self.logger.warning("Playwright 安装失败")
                        self.browser_combo.setCurrentIndex(0)  # 切换回本地浏览器
                        return
                elif browser_type == "undetected":
                    if not self.browser_installer.install_undetected_chrome():
                        self.logger.warning("Undetected Chrome 安装失败")
                        self.browser_combo.setCurrentIndex(0)  # 切换回本地浏览器
                        return
            else:
                self.logger.info("用户取消浏览器安装")
                self.browser_combo.setCurrentIndex(0)  # 切换回本地浏览器
                return
        
        self.config.set('browser', 'type', browser_type)
        self.logger.info("浏览器设置已更新")
        QMessageBox.information(self, "提示", "浏览器设置已更新，将在下次生成 PDF 时生效。")

    def toggle_pause(self):
        """切换暂停/继续状态"""
        if hasattr(self, 'generator_thread'):
            if self.generator_thread.is_paused:
                self.generator_thread.resume()
                self.pause_btn.setText("暂停")
                self.logger.info("继续生成 PDF")
            else:
                self.generator_thread.pause()
                self.pause_btn.setText("继续")
                self.logger.info("暂停生成 PDF")
                
    def stop_generation(self):
        """停止生成过程"""
        if hasattr(self, 'generator_thread'):
            self.generator_thread.stop()
            self.generate_btn.setEnabled(True)
            self.pause_btn.setEnabled(False)
            self.stop_btn.setEnabled(False)
            self.logger.info("停止生成 PDF")
            
    def update_status(self, status):
        """更新状态显示"""
        self.status_label.setText(status)

if __name__ == '__main__':
    try:
        # 检查 stdout 是否存在并支持 reconfigure
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        # 如果不支持 reconfigure，使用其他方式设置编码
        elif sys.stdout is not None:
            if hasattr(sys.stdout, 'encoding'):
                sys.stdout.encoding = 'utf-8'
    except Exception as e:
        # 忽略编码设置错误，不影响主程序运行
        pass

    # 禁用 Chrome 日志
    os.environ['WDM_LOG_LEVEL'] = '0'
    os.environ['WDM_PRINT_FIRST_LINE'] = 'False'
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 