"""测试发邮件：给富力达发送订单表"""
import smtplib
import socket
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_env():
    env_path = os.path.join(SCRIPT_DIR, ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    key, val = key.strip(), val.strip()
                    if key not in os.environ:
                        os.environ[key] = val

_load_env()

SMTP_SERVER = "smtp.qq.com"
SENDER = os.environ["QQ_EMAIL"]
AUTH_CODE = os.environ["QQ_AUTH_CODE"]
RECEIVER = "auroraqjy@163.com"  # 富力达测试用
DATE_STR = "20260624"
TIMEOUT = 30
import logging
logging.basicConfig(level=logging.DEBUG)

ATTACHMENT = os.path.join(SCRIPT_DIR, "output", "20260616_第2次", "富力达.xlsx")

if not os.path.exists(ATTACHMENT):
    # 回退找最新目录
    output_dir = os.path.join(SCRIPT_DIR, "output")
    for d in sorted(os.listdir(output_dir), reverse=True):
        path = os.path.join(output_dir, d, "富力达.xlsx")
        if os.path.exists(path):
            ATTACHMENT = path
            break

print(f"发件: {SENDER}")
print(f"收件: {RECEIVER}")
print(f"附件: {ATTACHMENT}")
sys.stdout.flush()

# 构造邮件
msg = MIMEMultipart()
msg["From"] = SENDER
msg["To"] = RECEIVER
msg["Subject"] = f"壳牌订单调度表_{DATE_STR}"
msg.attach(MIMEText("请查收今日订单，附件为贵司配送明细。", "plain", "utf-8"))

with open(ATTACHMENT, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header("Content-Disposition", "attachment", filename=os.path.basename(ATTACHMENT))
msg.attach(part)

# 尝试发送：先 465 SSL，再 587 STARTTLS
for port, use_ssl in [(465, True), (587, False)]:
    try:
        print(f"\n尝试连接 {SMTP_SERVER}:{port} ...")
        sys.stdout.flush()
        if use_ssl:
            server = smtplib.SMTP_SSL(SMTP_SERVER, port, timeout=TIMEOUT)
        else:
            server = smtplib.SMTP(SMTP_SERVER, port, timeout=TIMEOUT)
            server.set_debuglevel(2)
            print("已连接，启用 TLS ...")
            sys.stdout.flush()
            server.starttls()
            server.ehlo_or_helo_if_needed()
        print("已连接，正在登录...")
        sys.stdout.flush()
        server.login(SENDER, AUTH_CODE)
        print("登录成功，正在发送...")
        sys.stdout.flush()
        server.sendmail(SENDER, [RECEIVER], msg.as_string())
        server.quit()
        print(f"\n✅ 发送成功！（端口 {port}）请检查收件箱。")
        sys.stdout.flush()
        break
    except (socket.timeout, TimeoutError) as e:
        print(f"端口 {port} 连接超时")
    except Exception as e:
        print(f"端口 {port} 失败: {e}")
else:
    print("\n❌ 所有端口均失败。请检查：1) 网络是否可访问 smtp.163.com  2) 授权码是否正确")
