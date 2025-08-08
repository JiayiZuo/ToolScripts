import os, logging

# 日志文件路径
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)
SUCCESS_LOG = os.path.join(LOG_DIR, "email_success.log")
ERROR_LOG = os.path.join(LOG_DIR, "email_error.log")

def setup_logger():
    # 确保日志目录存在
    os.makedirs(LOG_DIR, exist_ok=True)

    """配置日志记录器"""
    logger = logging.getLogger("email_api")
    logger.setLevel(logging.DEBUG)

    # 成功日志处理器
    success_handler = logging.FileHandler(SUCCESS_LOG, encoding='utf-8')
    success_handler.setLevel(logging.INFO)
    success_format = logging.Formatter('%(asctime)s - %(message)s')
    success_handler.setFormatter(success_format)

    # 错误日志处理器
    error_handler = logging.FileHandler(ERROR_LOG, encoding='utf-8')
    error_handler.setLevel(logging.ERROR)
    error_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    error_handler.setFormatter(error_format)

    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    logger.addHandler(success_handler)
    logger.addHandler(error_handler)
    logger.addHandler(console_handler)

    return logger

logger = setup_logger()

