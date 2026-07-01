"""
SDCC 自动导出订单脚本
- 自动打开 Chrome，导航到 SDCC 并进行登录
- 自动执行：订单管理 → 搜索 → 导出
可作为模块被 run_all.py 导入调用，也可单独运行。
"""

import time
import os
import re
import glob
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException


def _get_script_dir():
	"""兼容 exe 打包：PyInstaller 的 sys._MEIPASS 或脚本目录"""
	if getattr(sys, 'frozen', False):
		return os.path.dirname(sys.executable)
	return os.path.dirname(os.path.abspath(__file__))


def _load_env():
	"""从 .env 文件加载配置（仅设置未定义的环境变量）"""
	env_path = os.path.join(_get_script_dir(), ".env")
	if os.path.exists(env_path):
		with open(env_path, "r", encoding="utf-8") as f:
			for line in f:
				line = line.strip()
				if line and not line.startswith("#") and "=" in line:
					key, val = line.split("=", 1)
					key, val = key.strip(), val.strip()
					if key not in os.environ:
						os.environ[key] = val


def run_sdcc_export(sdcc_date_str=None):
	"""
	执行 SDCC 自动导出。返回下载的文件路径，失败返回 None。
	sdcc_date_str 格式 YYYY-MM-DD，默认取环境变量 SDCC_DATE 或今天。
	"""
	_load_env()
	SCRIPT_DIR = _get_script_dir()

	URL = "https://tms.i.sinotrans.com/sdccweb/manage/welcompage"
	PROJECT_NAME = os.environ.get("SDCC_PROJECT", "壳牌天津8222")
	USERNAME = os.environ["SDCC_USERNAME"]
	PASSWORD = os.environ["SDCC_PASSWORD"]

	if sdcc_date_str is None:
		sdcc_date_str = os.environ.get("SDCC_DATE", datetime.now().strftime("%Y-%m-%d"))
	sdcc_date = datetime.strptime(sdcc_date_str, "%Y-%m-%d")
	sdcc_date_compact = sdcc_date.strftime("%Y%m%d")
	print(f"SDCC 导出日期: {sdcc_date_str}")

	DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "data", sdcc_date_compact)
	os.makedirs(DOWNLOAD_DIR, exist_ok=True)

	# 今天已经导出过了，跳过
	target_file = os.path.join(DOWNLOAD_DIR, f"download_{sdcc_date_compact}.xlsx")
	if os.path.exists(target_file):
		print(f"✅ 今天已导出: {target_file}")
		return target_file

	options = webdriver.ChromeOptions()
	options.add_argument("--log-level=3")
	options.add_experimental_option("excludeSwitches", ["enable-logging"])
	options.add_experimental_option("prefs", {
		"download.default_directory": DOWNLOAD_DIR,
		"download.prompt_for_download": False,
		"download.directory_upgrade": True,
	})
	driver = webdriver.Chrome(options=options)
	driver.maximize_window()
	driver.get(URL)

	print("=" * 60)
	print("  正在自动登录 SDCC...")
	print("=" * 60)
	time.sleep(5)
	main_window = driver.current_window_handle
	print(f"  当前 URL: {driver.current_url}")

	def close_popups():
		for handle in driver.window_handles:
			if handle != main_window:
				driver.switch_to.window(handle)
				print(f"  关闭标签页: {driver.title}")
				driver.close()
		driver.switch_to.window(main_window)
		try:
			btn = driver.find_element(By.CSS_SELECTOR, ".el-dialog__headerbtn")
			btn.click()
			print("  关闭了弹窗")
			time.sleep(1)
		except NoSuchElementException:
			pass

	def is_on_main_page():
		try:
			driver.find_element(By.XPATH, "//*[contains(text(),'我的工作台') or contains(text(),'订单中心')]")
			return True
		except NoSuchElementException:
			return False

	close_popups()

	if is_on_main_page():
		print("  检测到已在主界面，跳过登录")
	else:
		print("  点击「登录」进入登录页面...")
		try:
			driver.find_element(By.XPATH, "//span[text()='登录']/parent::*").click()
		except NoSuchElementException:
			driver.find_element(By.XPATH, "//*[text()='登录']").click()
		time.sleep(3)
		close_popups()
		print(f"  当前 URL: {driver.current_url}")

		try:
			internal_tab = driver.find_element(By.XPATH, "//div[@id='tab-IAM' or contains(text(),'内部用户')]")
			internal_tab.click()
			print("  选中「内部用户」")
			time.sleep(2)
		except NoSuchElementException:
			pass

		print("  等待登录 iframe 加载...")
		time.sleep(3)
		try:
			iframe = driver.find_element(By.ID, "iam_iframe_sdk")
			driver.switch_to.frame(iframe)
			print("  已进入登录 iframe")
		except NoSuchElementException:
			print("  ⚠ 未找到登录 iframe，尝试直接操作...")

		time.sleep(2)
		try:
			u = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='手机号码'], input[placeholder*='用户编号']")
			u.clear()
			u.send_keys(USERNAME)
			print(f"  用户名: {USERNAME}")

			p = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='请输入密码']")
			p.clear()
			p.send_keys(PASSWORD)
			print("  密码已输入")

			driver.find_element(By.CSS_SELECTOR, "button.login-button-click").click()
			print("  点击「登录」")

			driver.switch_to.default_content()
			time.sleep(5)
			close_popups()
			print(f"  登录后 URL: {driver.current_url}")

			if is_on_main_page():
				print("  登录成功！")
			else:
				print("  ⚠ 可能未登录成功，请检查后按 Enter 继续...")
				input()
		except NoSuchElementException:
			driver.switch_to.default_content()
			print("  ⚠ 未找到登录表单，请手动登录后按 Enter 继续...")
			input()
			close_popups()

	print("  登录阶段完成！")
	print("=" * 60)

	# [1/5] 检查导航
	print("\n[1/5] 检查导航...")
	time.sleep(2)
	try:
		workbench = driver.find_element(By.XPATH, "//*[contains(text(),'我的工作台')]")
		workbench.click()
		print("  点击了「我的工作台」")
		time.sleep(3)
	except NoSuchElementException:
		print("  已在首页，跳过")

	# [2/5] 进入订单管理
	print("[2/5] 进入订单管理...")
	time.sleep(2)
	try:
		order_center = driver.find_element(By.XPATH, "//*[contains(text(),'订单中心')]")
		order_center.click()
		time.sleep(1)
	except NoSuchElementException:
		pass

	try:
		order_mgmt = driver.find_element(By.XPATH, "//*[contains(text(),'订单管理')]")
		order_mgmt.click()
		print("  点击了「订单管理」")
		time.sleep(3)
	except NoSuchElementException:
		print("  ⚠ 未找到「订单管理」，尝试其他方式...")

	# [3/5] 输入项目名称
	print("\n[3/5] 输入项目名称...")
	driver.execute_script(f"""
		var inputs = document.querySelectorAll('input[placeholder="请输入选择"]');
		var target = inputs[inputs.length - 1];
		var select = target.closest('.el-select');
		select.click();
		setTimeout(function() {{
			var inner = select.querySelector('.el-select__input');
			if (inner) {{
				inner.value = '{PROJECT_NAME}';
				inner.dispatchEvent(new Event('input', {{ bubbles: true }}));
				inner.dispatchEvent(new Event('change', {{ bubbles: true }}));
			}}
		}}, 300);
	""")
	print(f"  输入: {PROJECT_NAME}")
	time.sleep(3)
	try:
		option = driver.find_element(By.XPATH, "//span[contains(text(),'壳牌天津')]")
		option.click()
		print(f"  选中: {PROJECT_NAME}")
	except NoSuchElementException:
		print("  ⚠ 未弹出下拉选项，请手动点击下拉项后按 Enter")
		input()

	# [4/5] 设置日期范围 + 搜索
	print("\n[4/5] 设置日期并搜索...")
	print(f"  设置日期范围: {sdcc_date_str} → {sdcc_date_str}")

	range_input = driver.find_element(By.CSS_SELECTOR, "input.el-range-input")
	range_input.click()
	time.sleep(1)

	sdcc_y, sdcc_m, sdcc_d = sdcc_date.year, sdcc_date.month, sdcc_date.day

	headers = driver.find_elements(By.CSS_SELECTOR, ".el-date-range-picker__header")
	left_text = headers[0].text if len(headers) > 0 else ""
	right_text = headers[1].text if len(headers) > 1 else ""
	print(f"  左面板: {left_text}")
	if right_text:
		print(f"  右面板: {right_text}")

	target_panel_idx = None
	panels_info = []
	for i, text in enumerate([left_text, right_text]):
		if not text:
			continue
		m = re.match(r'(\d{4})\s*年\s*(\d{1,2})', text)
		if m:
			py, pm = int(m.group(1)), int(m.group(2))
			panels_info.append((i, py, pm))
			if py == sdcc_y and pm == sdcc_m:
				target_panel_idx = i
				break

	if target_panel_idx is None:
		if panels_info:
			cur_year, cur_month = panels_info[0][1], panels_info[0][2]
		else:
			cur_year, cur_month = sdcc_y, sdcc_m
		month_diff = (sdcc_y - cur_year) * 12 + (sdcc_m - cur_month)
		print(f"  目标: {sdcc_y}年{sdcc_m}月, 左面板相差 {month_diff} 个月")

		if month_diff > 0:
			arrows = driver.find_elements(By.CSS_SELECTOR, ".el-icon-arrow-right")
			next_btn = arrows[0]
			for _ in range(month_diff):
				driver.execute_script("arguments[0].click();", next_btn)
				time.sleep(0.15)
		elif month_diff < 0:
			arrows = driver.find_elements(By.CSS_SELECTOR, ".el-icon-arrow-left")
			prev_btn = arrows[0]
			for _ in range(-month_diff):
				driver.execute_script("arguments[0].click();", prev_btn)
				time.sleep(0.15)

		time.sleep(0.5)
		target_panel_idx = 0
		print(f"  翻月完成，目标面板: 左")

	date_tables = driver.find_elements(By.CSS_SELECTOR, ".el-date-table")
	if target_panel_idx is not None and target_panel_idx < len(date_tables):
		target_table = date_tables[target_panel_idx]
	else:
		target_table = date_tables[0]

	day_cells = target_table.find_elements(By.CSS_SELECTOR, "td.available")
	target_cell = None
	for cell in day_cells:
		span = cell.find_element(By.TAG_NAME, "span")
		try:
			if int(span.text) == sdcc_d:
				target_cell = cell
				break
		except ValueError:
			continue

	if target_cell:
		target_cell.click()
		print(f"  开始日期 → {sdcc_date_str}")
		time.sleep(0.3)
		target_cell.click()
		print(f"  结束日期 → {sdcc_date_str}")
		time.sleep(0.5)

		confirm_btn = driver.find_element(By.XPATH, "//div[contains(@class,'el-picker-panel__footer')]//button[contains(.,'确定')]")
		confirm_btn.click()
		print("  点击「确定」")
		time.sleep(0.5)
	else:
		print(f"  ⚠ 未找到日期 {sdcc_d} 的单元格")

	time.sleep(1)
	print("  点击搜索...")
	try:
		search_btn = driver.find_element(By.XPATH, "//button//span[text()='搜索']")
		search_btn.click()
		print("  已点击「搜索」")
	except NoSuchElementException:
		try:
			search_btn = driver.find_element(By.XPATH, "//button[contains(text(),'搜索')]")
			search_btn.click()
			print("  已点击「搜索」")
		except NoSuchElementException:
			print("  ⚠ 未找到搜索按钮，请手动点击后按 Enter")
			input()

	time.sleep(5)

	# [5/5] 导出
	print("[5/5] 导出订单明细...")
	time.sleep(2)

	driver.execute_script("""
		var popover = document.querySelector('.el-popover[id]');
		if (popover) {
			var trigger = document.querySelector('[aria-describedby="' + popover.id + '"]');
			if (!trigger) trigger = popover.previousElementSibling;
			if (!trigger) {
				var all = document.querySelectorAll('[aria-describedby]');
				for (var i = 0; i < all.length; i++) {
					if (all[i].getAttribute('aria-describedby') === popover.id) {
						trigger = all[i]; break;
					}
				}
			}
			if (trigger) trigger.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
		}
	""")
	time.sleep(1)

	driver.execute_script("""
		var btns = document.querySelectorAll('.sdccDropDiv button');
		for (var i = 0; i < btns.length; i++) {
			if (btns[i].textContent.indexOf('订单明细记录') !== -1) {
				btns[i].click(); break;
			}
		}
	""")
	print("  点击「订单明细记录」")

	sdcc_file = None
	print("  等待下载完成...")
	start = time.time()
	before_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
	while time.time() - start < 60:
		crdownload = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
		if not crdownload:
			after_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
			new_files = after_files - before_files
			if new_files:
				sdcc_file = max(new_files, key=os.path.getctime)
				print(f"  下载完成: {os.path.basename(sdcc_file)}")
				break
		time.sleep(1)

	if sdcc_file:
		target_name = f"download_{sdcc_date_compact}.xlsx"
		target_path = os.path.join(DOWNLOAD_DIR, target_name)
		if sdcc_file != target_path:
			os.rename(sdcc_file, target_path)
			print(f"  已重命名: {os.path.basename(sdcc_file)} → {target_name}")
		sdcc_file = target_path
	else:
		print("  ⚠ 下载超时，请手动检查")

	driver.quit()

	if sdcc_file:
		print(f"\n✅ SDCC 导出完成: {os.path.basename(sdcc_file)}")
		return sdcc_file
	else:
		print("\n⚠ SDCC 导出可能未完成")
		return None


if __name__ == "__main__":
	result = run_sdcc_export()
	if result:
		print(f"\n导出成功: {result}")
	else:
		print("\n导出失败")
		sys.exit(1)
