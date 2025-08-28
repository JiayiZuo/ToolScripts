import string
import secrets

# 生成PDF随机密码
def generate_password():
    alphabet = string.ascii_letters + string.digits
    while True:
        password = ''.join(secrets.choice(alphabet) for i in range(10))
        if (any(c.islower() for c in password)
                and any(c.isupper() for c in password)
                and sum(c.isdigit() for c in password) >= 3):
            break
    return password

# 校验邮箱是否包含姓名
def check_email_name(name, email):
    if name.lower() in email.lower():
        return True
    return False