"""
SDCC 自动导出订单脚本
- 自动打开 Chrome，导航到 SDCC 并进行登录
- 自动执行：订单管理 → 搜索 → 导出
"""

import time
import os
import re
import glob
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def load_env():
    """从 .env 文件加载配置（仅设置未定义的环境变量）"""
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


load_env()

URL = "https://tms.i.sinotrans.com/sdccweb/manage/welcompage"
PROJECT_NAME = os.environ.get("SDCC_PROJECT", "壳牌天津8222")
USERNAME = os.environ["SDCC_USERNAME"]
PASSWORD = os.environ["SDCC_PASSWORD"]

DOWNLOAD_DIR = SCRIPT_DIR  # 下载到项目目录

# 可通过环境变量 SDCC_DATE 指定日期（格式 YYYY-MM-DD），默认今天
sdcc_date_str = os.environ.get("SDCC_DATE", datetime.now().strftime("%Y-%m-%d"))
sdcc_date = datetime.strptime(sdcc_date_str, "%Y-%m-%d")
sdcc_date_compact = sdcc_date.strftime("%Y%m%d")
print(f"SDCC 导出日期: {sdcc_date_str}")

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

# ============================================================
# 自动登录
# ============================================================
print("=" * 60)
print("  正在自动登录 SDCC...")
print("=" * 60)
time.sleep(5)
main_window = driver.current_window_handle

print(f"  当前 URL: {driver.current_url}")

# --- 辅助：关闭弹窗和多余标签页 ---
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

# --- 辅助：是否看到主界面 ---
def is_on_main_page():
    try:
        driver.find_element(By.XPATH, "//*[contains(text(),'我的工作台') or contains(text(),'订单中心')]")
        return True
    except NoSuchElementException:
        return False

close_popups()

# --- Step A: 如果已在主界面，无需登录 ---
if is_on_main_page():
    print("  检测到已在主界面，跳过登录")
else:
    # --- Step B: 点击「登录」进入登录页 ---
    print("  点击「登录」进入登录页面...")
    try:
        driver.find_element(By.XPATH, "//span[text()='登录']/parent::*").click()
    except NoSuchElementException:
        driver.find_element(By.XPATH, "//*[text()='登录']").click()
    time.sleep(3)
    close_popups()
    print(f"  当前 URL: {driver.current_url}")

    # --- Step C: 确保选中「内部用户」tab ---
    try:
        internal_tab = driver.find_element(By.XPATH, "//div[@id='tab-IAM' or contains(text(),'内部用户')]")
        internal_tab.click()
        print("  选中「内部用户」")
        time.sleep(2)
    except NoSuchElementException:
        pass

    # --- Step D: 切换到 iframe ---
    print("  等待登录 iframe 加载...")
    time.sleep(3)
    try:
        iframe = driver.find_element(By.ID, "iam_iframe_sdk")
        driver.switch_to.frame(iframe)
        print("  已进入登录 iframe")
    except NoSuchElementException:
        print("  ⚠ 未找到登录 iframe，尝试直接操作...")

    # --- Step E: 直接填写用户名密码（表单已在 iframe 中可见） ---
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

# ============================================================
# 第一步：检查并进入"我的工作台"
# ============================================================
print("\n[1/5] 检查导航...")
time.sleep(2)
try:
    workbench = driver.find_element(By.XPATH, "//*[contains(text(),'我的工作台')]")
    workbench.click()
    print("  点击了「我的工作台」")
    time.sleep(3)
except NoSuchElementException:
    print("  已在首页，跳过")

# ============================================================
# 第二步：点击左侧菜单 - 订单中心 > 订单管理
# ============================================================
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

# ============================================================
# 第三步：输入项目名称（JS 操作 Element UI el-select）
# ============================================================
print("\n[3/5] 输入项目名称...")
driver.execute_script(f"""
    // 找到 placeholder='请输入选择' 的 input，定位其 el-select 容器
    var inputs = document.querySelectorAll('input[placeholder="请输入选择"]');
    var target = inputs[inputs.length - 1];
    var select = target.closest('.el-select');
    // 点击容器打开下拉
    select.click();
    // 等 Vue 渲染后，向该容器内的 el-select__input 输入文字
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
# 点击下拉选项
try:
    option = driver.find_element(By.XPATH, "//span[contains(text(),'壳牌天津')]")
    option.click()
    print(f"  选中: {PROJECT_NAME}")
except NoSuchElementException:
    print("  ⚠ 未弹出下拉选项，请手动点击下拉项后按 Enter")
    input()

# ============================================================
# 第四步：设置日期范围 + 搜索
# ============================================================
print("\n[4/5] 设置日期并搜索...")

# 日期范围：用日历控件选取指定日期
print(f"  设置日期范围: {sdcc_date_str} → {sdcc_date_str}")

# 打开日期范围选择器
range_input = driver.find_element(By.CSS_SELECTOR, "input.el-range-input")
range_input.click()
time.sleep(1)

# 解析目标日期
sdcc_y, sdcc_m, sdcc_d = sdcc_date.year, sdcc_date.month, sdcc_date.day

# 读取左面板当前显示的月份
left_month_text = driver.find_element(By.CSS_SELECTOR, ".el-date-range-picker__header:first-child").text
print(f"  日历当前: {left_month_text}")

# 提取左面板的年份和月份
match = re.match(r'(\d{4})\s*年\s*(\d{1,2})', left_month_text)
if match:
    cur_year, cur_month = int(match.group(1)), int(match.group(2))
    # 计算要翻多少个月
    month_diff = (sdcc_y - cur_year) * 12 + (sdcc_m - cur_month)
    print(f"  目标: {sdcc_y}年{sdcc_m}月, 相差 {month_diff} 个月")

    # 翻月份
    if month_diff > 0:
        next_btn = driver.find_element(By.CSS_SELECTOR, ".el-icon-arrow-right")
        for _ in range(month_diff):
            next_btn.click()
            time.sleep(0.1)
    elif month_diff < 0:
        prev_btn = driver.find_element(By.CSS_SELECTOR, ".el-icon-arrow-left")
        for _ in range(-month_diff):
            prev_btn.click()
            time.sleep(0.1)

    time.sleep(0.5)

    # 点击目标日期单元格（在左面板中）
    # td.available 且 span 文本等于目标日
    day_cells = driver.find_elements(By.CSS_SELECTOR, "td.available")
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

        # 点击确定
        confirm_btn = driver.find_element(By.XPATH, "//div[contains(@class,'el-picker-panel__footer')]//button[contains(.,'确定')]")
        confirm_btn.click()
        print("  点击「确定」")
        time.sleep(0.5)
    else:
        print(f"  ⚠ 未找到日期 {sdcc_d} 的单元格，回退到直接输入")
        raise NoSuchElementException("day not found")
else:
    raise NoSuchElementException("month parse failed")

# 点击搜索按钮
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

# ============================================================
# 第五步：导出（悬停图标 → 弹出菜单 → 点击「订单明细记录」）
# ============================================================
print("[5/5] 导出订单明细...")
time.sleep(2)

# 方案1: 用 JS 悬停触发元素，让 popover 显示
driver.execute_script("""
    // 找 el-popover 的触发元素（aria-describedby 指向 popover id）
    var popover = document.querySelector('.el-popover[id]');
    if (popover) {
        var popoverId = '#' + popover.id;
        var trigger = document.querySelector('[aria-describedby="' + popover.id + '"]');
        if (!trigger) {
            // 可能 trigger 在 popover 前面，找同级元素
            trigger = popover.previousElementSibling;
        }
        if (!trigger) {
            // 通过 v-popover 找，找所有带 popover 引用的元素
            var all = document.querySelectorAll('[aria-describedby]');
            for (var i = 0; i < all.length; i++) {
                if (all[i].getAttribute('aria-describedby') === popover.id) {
                    trigger = all[i];
                    break;
                }
            }
        }
        if (trigger) {
            trigger.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
        }
    }
""")
time.sleep(1)

# 方案2: 直接用 JS 点击隐藏的「订单明细记录」按钮
driver.execute_script("""
    var btns = document.querySelectorAll('.sdccDropDiv button');
    for (var i = 0; i < btns.length; i++) {
        if (btns[i].textContent.indexOf('订单明细记录') !== -1) {
            btns[i].click();
            break;
        }
    }
""")
print("  点击「订单明细记录」")

# 等待下载完成
sdcc_file = None
print("  等待下载完成...")
start = time.time()
before_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
while time.time() - start < 60:
    # 检查是否还有正在下载的文件
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
    # 重命名为 download_YYYYMMDD.xlsx 以匹配 order_merge 的文件识别规则
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
else:
    print("\n⚠ SDCC 导出可能未完成")
