"""
壳牌订单调度 - 全流程自动化
1. 下载客户邮件附件
2. 从客户表提取最新子表日期，推算 SDCC 导出日期（子表日期-1天）
3. SDCC 导出订单
4. 合并处理 + 发邮件给各承运商
"""

import subprocess
import sys
import os
import re
import glob
from datetime import datetime, timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENV_PYTHON = os.path.join(SCRIPT_DIR, "venv", "bin", "python3")


def run_step(name, script, env=None):
    print("\n" + "=" * 60)
    print(f"  {name}")
    print("=" * 60)
    run_env = os.environ.copy()
    if env:
        run_env.update(env)
    result = subprocess.run(
        [VENV_PYTHON, os.path.join(SCRIPT_DIR, script)],
        cwd=SCRIPT_DIR,
        env=run_env,
    )
    if result.returncode != 0:
        print(f"\n❌ {name} 失败 (exit code {result.returncode})")
        return False
    return True


def find_customer_file():
    """找到项目目录下的客户文件（非 SDCC download_ 开头的 xlsx）"""
    files = glob.glob(os.path.join(SCRIPT_DIR, "*.xlsx"))
    for f in sorted(files, key=os.path.getmtime, reverse=True):
        basename = os.path.basename(f)
        if not basename.startswith("download_") and not basename.startswith("~$"):
            return f
    return None


def get_latest_subtable_date(customer_file):
    """读取客户文件，返回最新子表日期 (month, day)"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(customer_file, read_only=True)
        sheets = wb.sheetnames
        wb.close()

        max_md = (0, 0)
        for name in sheets:
            m = re.match(r'^(\d{2})(\d{2})$', name.strip())
            if m:
                mm, dd = int(m.group(1)), int(m.group(2))
                if 1 <= mm <= 12 and 1 <= dd <= 31:
                    if (mm, dd) > max_md:
                        max_md = (mm, dd)

        if max_md != (0, 0):
            print(f"  客户表最新子表: {max_md[0]:02d}{max_md[1]:02d}")
            return max_md
    except Exception as e:
        print(f"  ⚠ 读取客户文件失败: {e}")
    return None


def calc_sdcc_date(sub_mm, sub_dd):
    """子表日期 - 1天 = SDCC 导出日期"""
    today = datetime.now()
    # 假设子表日期在当年或下一年（跨年场景）
    sub_date = datetime(today.year, sub_mm, sub_dd)
    if sub_date > today + timedelta(days=60):
        # 子表日期太远，可能是去年
        sub_date = datetime(today.year - 1, sub_mm, sub_dd)
    sdcc_date = sub_date - timedelta(days=1)
    return sdcc_date.strftime("%Y-%m-%d")


if __name__ == "__main__":
    print("=" * 60)
    print("  壳牌订单调度 - 全流程自动化")
    print("=" * 60)

    # Step 1: 下载客户邮件附件
    if not run_step("Step 1/4: 下载客户邮件附件", "email_download.py"):
        print("邮件下载失败，尝试使用已有客户文件...")

    # Step 2: 推算 SDCC 导出日期
    customer_file = find_customer_file()
    sdcc_env = {}
    if customer_file:
        print(f"\n客户文件: {os.path.basename(customer_file)}")
        md = get_latest_subtable_date(customer_file)
        if md:
            sdcc_date = calc_sdcc_date(md[0], md[1])
            sdcc_env["SDCC_DATE"] = sdcc_date
            print(f"SDCC 导出日期（子表-1天）: {sdcc_date}")
        else:
            print("未能提取子表日期，SDCC 将使用今天")
    else:
        print("未找到客户文件，SDCC 将使用今天")

    # Step 3: SDCC 导出
    if not run_step("Step 3/4: SDCC 导出订单", "sdcc_export.py", env=sdcc_env):
        print("SDCC 导出失败，流程终止。")
        sys.exit(1)

    # Step 4: 合并处理 + 发邮件
    if not run_step("Step 4/4: 订单合并 + 发送邮件", "order_merge.py"):
        print("订单处理失败，请检查。")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  ✅ 全流程完成！")
    print("=" * 60)
