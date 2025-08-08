import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import os, uuid
from log_config import logger
from flask import Blueprint, request, jsonify
from dotenv import load_dotenv
from common import code, message

email_bp = Blueprint('email_api', __name__, url_prefix='/api/email')
load_dotenv()

# 从环境变量获取配置
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.163.com")  # 默认值
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))  # 转换为整数
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
TEMPLATE_FILE = os.getcwd() + "/utils/salary_email.html"
SUBJECT_TEMPLATE = "{}/{}/{}明细"


def send_emails(excel_path):
    # 获取当前日期（日、月、年）
    today = datetime.now()
    day = f"{today.day:02d}"
    month = f"{today.month:02d}"
    year = today.year

    # 生成主题
    subject = SUBJECT_TEMPLATE.format(day, month, year)

    # 读取Excel数据
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        logger.error(f"读取Excel失败: {e}")
        return False, str(e)

    # 加载邮件模板
    try:
        with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
            template = f.read()
    except Exception as e:
        logger.error(f"加载模板失败: {e}")
        return False, str(e)

    # 连接SMTP服务器
    try:
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
    except Exception as e:
        logger.error(f"SMTP连接失败: {e}")
        return False, str(e)

    # 发送邮件
    success_count = 0
    error_details = []
    for index, row in df.iterrows():
        try:
            # 检查必要字段
            if '邮箱' not in row or not row['邮箱']:
                raise ValueError("缺少邮箱地址")

            # 生成个性化邮件内容
            personalized_content = template
            for col in df.columns:
                placeholder = f"{{{{{col}}}}}"
                personalized_content = personalized_content.replace(placeholder, str(row[col]))

            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = row['邮箱']
            msg['Subject'] = subject
            msg.attach(MIMEText(personalized_content, "html"))

            # 发送邮件
            server.sendmail(SENDER_EMAIL, [row['邮箱']], msg.as_string())

            # 记录成功
            recipient_info = f"{row.get('姓名', '未知')} <{row['邮箱']}>"
            success_count += 1
        except Exception as e:
            # 记录错误
            error_msg = f"发送给 {row.get('姓名', '未知')} 失败: {str(e)}"
            logger.error(error_msg)
            error_details.append({
                "recipient": row.get('姓名', '未知'),
                "email": row.get('邮箱', '无邮箱'),
                "error": str(e)
            })

    # 关闭SMTP连接
    server.quit()

    # 返回结果
    total = len(df)
    result = {
        "success": success_count,
        "total": total,
        "error_count": total - success_count,
        "error_details": error_details,
        "subject": subject
    }

    if success_count == total:
        logger.info(f"所有邮件发送成功! 总计: {success_count}封")
    else:
        logger.warning(f"邮件发送完成! 成功: {success_count}/{total} 封")

    return True, result

@email_bp.route('/send-salary-emails', methods=['POST'])
def handle_send_emails():
    # 生成唯一ID用于日志跟踪
    request_id = str(uuid.uuid4())

    # 检查请求中是否有文件
    if 'excel' not in request.files not in request.files:
        logger.error(f"[{request_id}] 缺少必要文件")
        return jsonify({
            "status": "error",
            "message": "请提供excel和template文件"
        }), 400

    excel_file = request.files['excel']

    # 发送邮件
    success, result = send_emails(excel_file)

    # 返回结果
    if success:
        logger.info(f"[{request_id}] 请求处理成功")
        return jsonify({
            "data": {
                "request_id": request_id,
                **result
            },
            "message": message.SUCCESS,
            "code": code.SUCCESS
        })
    else:
        logger.error(f"[{request_id}] 请求处理失败")
        return jsonify({
            "data": {
                "request_id": request_id,
                "result": result
            },
            "message": message.SALARY_EMAIL_FAILED,
            "code": code.SALARY_EMAIL_FAIL
        }), 500