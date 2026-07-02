"""
壳牌订单调度 - 全流程自动化
1. 下载客户邮件附件
2. 从客户表提取最新子表日期，推算 SDCC 导出日期（子表日期-1天）
3. SDCC 导出订单
4. 合并处理 + 发邮件给各承运商
"""

import sys
import os
import re
import glob
import shutil
from datetime import datetime, timedelta

# 兼容 exe 打包
if getattr(sys, 'frozen', False):
	SCRIPT_DIR = os.path.dirname(sys.executable)
else:
	SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

from email_download import download_customer_attachment
from sdcc_export import run_sdcc_export
from order_merge import main as order_merge_main


def find_customer_file():
	"""找到最新的客户文件（非 SDCC download_ 开头的 xlsx），含 data/ 子目录"""
	candidates = []
	for f in glob.glob(os.path.join(SCRIPT_DIR, "*.xlsx")):
		basename = os.path.basename(f)
		if not basename.startswith("download_") and not basename.startswith("~$"):
			candidates.append(f)
	data_dir = os.path.join(SCRIPT_DIR, "data")
	if os.path.isdir(data_dir):
		cutoff = datetime.now() - timedelta(days=7)
		for sub in sorted(os.listdir(data_dir), reverse=True):
			sub_path = os.path.join(data_dir, sub)
			if os.path.isdir(sub_path):
				m = re.match(r'^(\d{4})(\d{2})(\d{2})$', sub)
				if m:
					sub_date = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
					if sub_date < cutoff:
						continue
				for f in glob.glob(os.path.join(sub_path, "*.xlsx")):
					basename = os.path.basename(f)
					if not basename.startswith("download_") and not basename.startswith("~$"):
						candidates.append(f)
	if not candidates:
		return None
	return max(candidates, key=os.path.getmtime)


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
	sub_date = datetime(today.year, sub_mm, sub_dd)
	if sub_date > today + timedelta(days=60):
		sub_date = datetime(today.year - 1, sub_mm, sub_dd)
	sdcc_date = sub_date - timedelta(days=1)
	return sdcc_date.strftime("%Y-%m-%d")


if __name__ == "__main__":
	print("=" * 60)
	print("  壳牌订单调度 - 全流程自动化")
	print("=" * 60)

	# 清理昨天的 output
	yesterday = datetime.now() - timedelta(days=1)
	yesterday_dir = os.path.join(SCRIPT_DIR, "output", yesterday.strftime("%Y%m%d"))
	if os.path.exists(yesterday_dir):
		shutil.rmtree(yesterday_dir)
		print(f"🧹 已清理昨日输出: {yesterday.strftime('%Y%m%d')}")

	# 检查今天是否已完成
	today_str = datetime.now().strftime("%Y%m%d")
	today_output = os.path.join(SCRIPT_DIR, "output", today_str)
	if os.path.exists(os.path.join(today_output, "调度总表.xlsx")):
		print(f"\n✅ 今天({today_str})已完成，无需重复执行。")
		sys.exit(0)

	# Step 1: 下载客户邮件附件
	print("\n" + "=" * 60)
	print("  Step 1/4: 下载客户邮件附件")
	print("=" * 60)
	customer_downloaded = download_customer_attachment()
	if not customer_downloaded:
		print("\n❌ 今天未收到新的客户发货计划邮件，流程终止。")
		sys.exit(0)

	# Step 2: 推算 SDCC 导出日期
	print("\n" + "=" * 60)
	print("  Step 2/4: 推算 SDCC 导出日期")
	print("=" * 60)
	customer_file = find_customer_file()
	sdcc_date_str = None
	if customer_file:
		print(f"客户文件: {os.path.basename(customer_file)}")
		md = get_latest_subtable_date(customer_file)
		if md:
			sdcc_date_str = calc_sdcc_date(md[0], md[1])
			os.environ["SDCC_DATE"] = sdcc_date_str
			print(f"SDCC 导出日期（子表-1天）: {sdcc_date_str}")
		else:
			print("未能提取子表日期，SDCC 将使用今天")
	else:
		print("未找到客户文件，SDCC 将使用今天")

	# Step 3: SDCC 导出
	print("\n" + "=" * 60)
	print("  Step 3/4: SDCC 导出订单")
	print("=" * 60)
	sdcc_result = run_sdcc_export(sdcc_date_str)
	if not sdcc_result:
		print("\n❌ SDCC 导出失败，流程终止。")
		sys.exit(1)

	# Step 4: 合并处理 + 发邮件
	print("\n" + "=" * 60)
	print("  Step 4/4: 订单合并 + 发送邮件")
	print("=" * 60)
	order_merge_main()

	print("\n" + "=" * 60)
	print("  ✅ 全流程完成！")
	print("=" * 60)
