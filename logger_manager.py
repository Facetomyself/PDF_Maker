import os
import sys
from loguru import logger
from datetime import datetime
from utils import resource_path

class LoggerManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LoggerManager, cls).__new__(cls)
            cls._instance._initialize_logger()
        return cls._instance
    
    def _initialize_logger(self):
        """初始化日志配置"""
        # 创建 logs 文件夹
        log_dir = resource_path('logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        # 生成日志文件名
        current_time = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(log_dir, f"pdf_maker_{current_time}.log")
        
        # 移除默认的处理器
        logger.remove()
        
        try:
            # 尝试添加控制台输出
            if sys.stderr is not None:
                logger.add(
                    sys.stderr,
                    format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
                    level="INFO"
                )
        except TypeError:
            # 如果控制台输出失败，忽略错误
            pass
        
        # 添加文件输出
        logger.add(
            log_file,
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            level="DEBUG",
            rotation="00:00",  # 每天午夜轮换
            retention="30 days",  # 保留30天
            encoding="utf-8",
            enqueue=True  # 启用异步写入
        )
        
    def get_logger(self):
        """获取日志记录器"""
        return logger 