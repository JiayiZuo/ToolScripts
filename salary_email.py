import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import os, uuid, tempfile
from log_config import logger
from flask import Blueprint, request, jsonify
from dotenv import load_dotenv
from common import code, message
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import matplotlib.font_manager as fm
from PyPDF2 import PdfReader, PdfWriter  # 添加PDF加密所需的库

# 尝试注册中文字体
try:
    # 查找系统中可用的中文字体
    chinese_fonts = [f.name for f in fm.fontManager.ttflist if any(
        word in f.name.lower() for word in ['chinese', 'china', 'simhei', 'simsun', 'msyh', 'pingfang'])]

    if chinese_fonts:
        # 使用第一个找到的中文字体
        chinese_font_path = fm.findfont(fm.FontProperties(family=chinese_fonts[0]))
        pdfmetrics.registerFont(TTFont('ChineseFont', chinese_font_path))
        logger.info(f"使用中文字体: {chinese_fonts[0]}, 路径: {chinese_font_path}")
    else:
        # 如果没有找到中文字体，尝试使用默认字体
        pdfmetrics.registerFont(TTFont('ChineseFont', 'arial.ttf'))
        logger.warning("未找到中文字体，使用默认字体")
except Exception as e:
    logger.error(f"字体注册失败: {e}")
    # 如果所有尝试都失败，使用ReportLab内置的CID字体
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    logger.info("使用内置CID字体: STSong-Light")

email_bp = Blueprint('email_api', __name__, url_prefix='/api/email')
load_dotenv()

# 从环境变量获取配置
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.163.com")  # 默认值
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))  # 转换为整数
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
LOGO_PATH = os.getenv("LOGO_PATH", os.getcwd() + "/utils/logo.png")  # 公司LOGO路径
TEMPLATE_FILE = os.getcwd() + "/utils/salary_email.html"
SUBJECT_TEMPLATE = "{}/{}/{}明细"
PDF_PASSWORD = os.getenv("PDF_PASSWORD", "123456")  # 添加PDF密码环境变量


def read_excel_data(excel_file):
    """
    从Excel文件读取指定字段的数据并格式化
    """
    try:
        # 读取Excel文件
        excel_file.seek(0)
        df = pd.read_excel(excel_file, engine='openpyxl')

        # 确保所需的列存在
        required_columns = ['基本薪金', 'TR_FEE', '月度奖金', '佣金', '其他', 'MPF', '总共']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要的列: {col}")

        # 格式化数据为两位小数
        for col in required_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)

        return df
    except Exception as e:
        logger.error(f"读取Excel文件时出错: {e}")
        return None


def encrypt_pdf(input_path, output_path, password):
    try:
        with open(input_path, "rb") as input_file:
            pdf_reader = PdfReader(input_file)
            pdf_writer = PdfWriter()

            # 复制所有页面
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

            # 加密PDF
            pdf_writer.encrypt(password)

            # 保存加密后的PDF
            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

        return True
    except Exception as e:
        logger.error(f"加密PDF失败: {e}")
        return False


def create_pdf(dataframe, pdf_path, logo_path):
    try:
        # 创建PDF文档
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)
        elements = []

        # 添加公司LOGO
        if os.path.exists(logo_path):
            logo = Image(logo_path, width=2 * inch, height=1 * inch)
            logo.hAlign = 'CENTER'
            elements.append(logo)
            elements.append(Spacer(1, 0.2 * inch))

        # 添加标题 - 使用中文字体
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontName='ChineseFont',
            fontSize=16,
            spaceAfter=30,
            alignment=1  # 居中
        )
        title = Paragraph("员工薪酬明细表", title_style)
        elements.append(title)
        elements.append(Spacer(1, 0.2 * inch))

        # 准备表格数据
        table_data = [['组成部分', '金额(元)']]  # 表头

        # 添加数据行
        row = dataframe.iloc[0]  # 获取第一行数据
        table_data.append(['基本薪金', f"{row['基本薪金']:.2f}"])
        table_data.append(['TR FEE', f"{row['TR_FEE']:.2f}"])
        table_data.append(['月度奖金', f"{row['月度奖金']:.2f}"])
        table_data.append(['佣金', f"{row['佣金']:.2f}"])
        table_data.append(['其他', f"{row['其他']:.2f}"])
        table_data.append(['MPF', f"{row['MPF']:.2f}"])
        table_data.append(['总共', f"{row['总共']:.2f}"])

        # 创建表格
        table = Table(table_data, colWidths=[2.5 * inch, 2.5 * inch])

        # 设置表格样式 - 使用中文字体
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'ChineseFont'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('FONTNAME', (0, 1), (-1, -1), 'ChineseFont'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])

        # 为总计行添加特殊样式
        table_style.add('BACKGROUND', (0, len(table_data) - 1), (-1, len(table_data) - 1), colors.lightgrey)
        table_style.add('FONTSIZE', (0, len(table_data) - 1), (-1, len(table_data) - 1), 14)
        table_style.add('FONTNAME', (0, len(table_data) - 1), (-1, len(table_data) - 1), 'ChineseFont')

        table.setStyle(table_style)
        elements.append(table)

        # 生成PDF
        doc.build(elements)
        logger.info(f"PDF文件已生成: {pdf_path}")
        return True
    except Exception as e:
        logger.error(f"生成PDF文件时出错: {e}")
        return False


def send_emails(excel_path):
    # 获取当前日期（日、月、年）
    today = datetime.now()
    day = f"{today.day:02d}"
    month = f"{today.month:02d}"
    year = today.year

    # 生成主题
    subject = SUBJECT_TEMPLATE.format(day, month, year)

    if not excel_path.filename.endswith(('.xlsx', '.xls')):
        return False, "只支持Excel文件(.xlsx, .xls)"

    # 读取Excel数据
    df = read_excel_data(excel_path)
    if df is None:
        return False, "读取Excel数据失败"

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

            # 创建临时PDF文件（未加密）
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_file:
                unencrypted_pdf_path = tmp_file.name

            # 为当前行创建PDF
            pdf_created = create_pdf(pd.DataFrame([row]), unencrypted_pdf_path, LOGO_PATH)
            if not pdf_created:
                raise ValueError("生成PDF文件失败")

            # 创建加密的PDF文件
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_file:
                encrypted_pdf_path = tmp_file.name

            # 加密PDF
            encrypt_success = encrypt_pdf(unencrypted_pdf_path, encrypted_pdf_path, PDF_PASSWORD)
            if not encrypt_success:
                raise ValueError("加密PDF文件失败")

            # 删除未加密的临时文件
            os.unlink(unencrypted_pdf_path)

            # 添加PDF密码提示
            password_note = f"<p><strong>请注意：</strong>PDF附件已加密，打开密码为：{PDF_PASSWORD}</p>"
            template = template.replace("</body>", f"{password_note}</body>")

            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = row['邮箱']
            msg['Subject'] = subject
            msg.attach(MIMEText(template, "html"))

            # 添加PDF附件
            with open(encrypted_pdf_path, "rb") as f:
                attach = MIMEApplication(f.read(), _subtype="pdf")
            attach.add_header('Content-Disposition', 'attachment', filename=f"薪酬明细_{day}{month}{year}.pdf")
            msg.attach(attach)

            # 发送邮件
            server.sendmail(SENDER_EMAIL, [row['邮箱']], msg.as_string())

            # 删除临时文件
            os.unlink(encrypted_pdf_path)

            # 记录成功
            recipient_info = f"{row.get('姓名', '未知')} <{row['邮箱']}>"
            success_count += 1
            logger.info(f"邮件发送成功: {recipient_info}")
        except Exception as e:
            # 记录错误
            error_msg = f"发送给 {row.get('姓名', '未知')} 失败: {str(e)}"
            logger.error(error_msg)
            error_details.append({
                "recipient": row.get('姓名', '未知'),
                "email": row.get('邮箱', '无邮箱'),
                "error": str(e)
            })

            # 确保临时文件被删除
            for path_var in ['unencrypted_pdf_path', 'encrypted_pdf_path']:
                if path_var in locals() and os.path.exists(locals()[path_var]):
                    try:
                        os.unlink(locals()[path_var])
                    except:
                        pass

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
    if 'excel' not in request.files:
        logger.error(f"[{request_id}] 缺少必要文件")
        return jsonify({
            "status": "error",
            "message": "请提供excel文件"
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