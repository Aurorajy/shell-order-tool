"""
从公司邮箱下载客户发货计划附件
"""

import imaplib
import email
from email.header import decode_header
import os
import re
import sys
import glob
from datetime import datetime, timedelta

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
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

IMAP_SERVER = os.environ.get("EMAIL_IMAP_SERVER", "imap.qq.com")
IMAP_PORT = int(os.environ.get("EMAIL_IMAP_PORT", "993"))
USERNAME = os.environ["EMAIL_USERNAME"]
AUTH_CODE = os.environ["EMAIL_PASSWORD"]


def download_customer_attachment(date_override=None):
    """
    下载今天收到的最新一封带附件的 Shell 发货计划邮件。
    不再匹配邮件标题中的日期，只看邮件接收时间是否为今天。
    保存到 data/YYYYMMDD/ 目录，返回下载的文件路径，失败返回 None。
    """
    today = date_override or datetime.now()
    date_dir = os.path.join(SCRIPT_DIR, "data", today.strftime("%Y%m%d"))
    os.makedirs(date_dir, exist_ok=True)

    # 今天已经下载过了，直接返回已有文件
    existing = glob.glob(os.path.join(date_dir, "*.xlsx")) + glob.glob(os.path.join(date_dir, "*.xls"))
    existing = [f for f in existing if not os.path.basename(f).startswith(("~$", ".~", "download_"))]
    if existing:
        latest = max(existing, key=os.path.getmtime)
        print(f"✅ 今天已有附件: {os.path.basename(latest)}")
        return latest

    # IMAP SINCE 日期格式: DD-Mon-YYYY，从昨天起搜（留一天余量）
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    yday = today - timedelta(days=1)
    since_yesterday = f"{yday.day:02d}-{months[yday.month-1]}-{yday.year}"

    print(f"连接 {IMAP_SERVER}:{IMAP_PORT} ...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(USERNAME, AUTH_CODE)
    print("登录成功")

    status, result = mail.select("INBOX")
    if status != "OK":
        print(f"❌ 无法选择收件箱: {result}")
        mail.logout()
        return None

    # 只搜从昨天起收到的邮件（留一天余量），大幅减少搜索范围
    status, data = mail.search(None, '(SUBJECT "Shell" SINCE {})'.format(since_yesterday))
    email_ids = data[0].split() if data[0] else []

    if not email_ids:
        print("❌ 未找到今天收到的 Shell 邮件")
        mail.logout()
        return None

    print(f"从昨天起收到 {len(email_ids)} 封含 'Shell' 的邮件")

    # 从最新到最旧遍历，取第一封匹配的（带附件）
    downloaded = None
    for eid in reversed(email_ids):
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])

        # 只处理今天收到的邮件
        msg_date_str = msg.get("Date", "")
        try:
            # 解析邮件日期，格式: "Mon, 3 Aug 2026 13:55:18 +0800"
            from email.utils import parsedate_to_datetime
            msg_date = parsedate_to_datetime(msg_date_str)
            if msg_date.date() != today.date():
                continue
        except Exception:
            pass  # 无法解析日期则不过滤

        subject, enc = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(enc or "utf-8", errors="replace")
        subject_clean = re.sub(r'^(转发[：:]|Fwd?:|RE:)\s*', '', subject, flags=re.IGNORECASE)

        if not ("Shell" in subject_clean and "发货计划" in subject_clean):
            continue

        print(f"  匹配: {subject_clean[:60]}")

        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                fname, enc = decode_header(filename)[0]
                if isinstance(fname, bytes):
                    fname = fname.decode(enc or "utf-8", errors="replace")
                if fname.lower().endswith((".xlsx", ".xls")):
                    save_path = os.path.join(date_dir, fname)
                    with open(save_path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    downloaded = save_path
                    print(f"  ✅ 下载附件: {fname}")
                    break
        if downloaded:
            break

    mail.logout()

    if downloaded:
        print(f"✅ 客户附件: {os.path.basename(downloaded)}")
    else:
        print("❌ 未找到今天带附件的 Shell 发货计划邮件")
    return downloaded


if __name__ == "__main__":
    result = download_customer_attachment()
    if result:
        print(f"\n下载完成: {result}")
    else:
        print("\n下载失败")
