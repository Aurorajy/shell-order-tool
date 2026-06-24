"""
从 QQ 邮箱下载客户发货计划附件
"""

import imaplib
import email
from email.header import decode_header
import os
import re
from datetime import datetime

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

IMAP_SERVER = "imap.qq.com"
IMAP_PORT = 993
USERNAME = os.environ["QQ_EMAIL"]
AUTH_CODE = os.environ["QQ_AUTH_CODE"]


def download_customer_attachment(date_override=None):
    """
    下载今天（或指定日期）的客户发货计划附件。
    返回下载的文件路径，失败返回 None。
    """
    today = date_override or datetime.now()
    date_str = f"{today.year}/{today.month}/{today.day}"
    date_str2 = f"{today.year}/{today.month:02d}/{today.day:02d}"

    print(f"连接 {IMAP_SERVER}:{IMAP_PORT} ...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(USERNAME, AUTH_CODE)
    print("登录成功")

    status, result = mail.select("INBOX")
    if status != "OK":
        print(f"❌ 无法选择收件箱: {result}")
        mail.logout()
        return None

    # IMAP 不支持中文搜索，用英文关键词搜
    status, data = mail.search(None, 'SUBJECT', 'Shell')
    email_ids = data[0].split() if data[0] else []

    if not email_ids:
        print("❌ 未找到相关邮件")
        mail.logout()
        return None

    print(f"找到 {len(email_ids)} 封含 'Shell' 的邮件")

    downloaded = None
    for eid in reversed(email_ids):
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        subject, enc = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(enc or "utf-8", errors="replace")
        subject_clean = re.sub(r'^(转发[：:]|Fwd?:|RE:)\s*', '', subject, flags=re.IGNORECASE)

        if not ("Shell" in subject_clean and "发货计划" in subject_clean):
            continue
        if not (date_str in subject_clean or date_str2 in subject_clean):
            continue

        print(f"  匹配: {subject_clean[:60]}")

        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                fname, enc = decode_header(filename)[0]
                if isinstance(fname, bytes):
                    fname = fname.decode(enc or "utf-8", errors="replace")
                if fname.lower().endswith((".xlsx", ".xls")):
                    save_path = os.path.join(SCRIPT_DIR, fname)
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
