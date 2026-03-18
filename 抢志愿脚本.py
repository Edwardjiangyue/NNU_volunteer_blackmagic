import time
import keyboard
import pyperclip
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException as TE
from selenium.common.exceptions import ElementClickInterceptedException as ECIE
from selenium.webdriver.chrome.service import Service
import yaml
import os
import re


# 测试表单地址：
# 金数据1：https://jsj.top/f/b4cinX
# 金数据2：https://zsmjxc1m.jsjform.com/f/irJ0VJ
# 问卷星：https://www.wjx.top/vm/YZjvTpC.aspx#
# 如果测试表单进不去，可能是链接失效了，可以自己创建一个表单

def geturl():
    """
    监听系统剪贴板，获取用户手动复制的表单链接（支持QQ复制后自动捕获）
    支持的链接格式：以http/https开头，且包含表单特征（jsj/wjx）
    :return: 符合要求的表单URL字符串
    """
    url_pattern = re.compile(r"https://[^\s\n]+")
    print("=== 请获取表单链接 ===")
    print("1. 打开QQ群，找到目标表单链接")
    print("2. 右键点击链接 → 选择【复制】（或按 Ctrl+C）")
    print("3. 脚本将自动捕获复制的链接...\n")
    last_clipboard = pyperclip.paste().strip()
    while True:
        current_clipboard = pyperclip.paste().strip()
        if current_clipboard != last_clipboard:
            match_result = url_pattern.search(current_clipboard)
            if match_result:
                form_url = match_result.group()
                print(f"✅ 成功捕获表单链接：{form_url}")
                return form_url
            print(f'❌ 复制的内容不包含有效表单链接，请确保您要访问的链接是以"https://"开头的完整链接')
            last_clipboard = current_clipboard
        time.sleep(0.3)


def explicit_find_xpath(driver, xpath):
    """通过XPath显式等待定位元素，确保稳定性"""
    try:
        WebDriverWait(driver, 1.5, 0.5).until(EC.presence_of_element_located(('xpath', xpath)))
        return driver.find_element('xpath', xpath)
    except TE:
        print(f"元素未找到：XPath={xpath}")
        return None


def explicit_click_xpath(driver, xpath):
    """通过XPath点击元素"""
    ele = explicit_find_xpath(driver, xpath)
    if not ele:
        return
    try:
        WebDriverWait(driver, 1.5, 0.5).until(EC.element_to_be_clickable(('xpath', xpath)))
        ele.click()
    except (ECIE, TE):
        driver.execute_script("arguments[0].click();", ele)


def auto_fill_generic(driver, configs):
    """
    严格按你的规则实现：
    1. 父容器XPath：//div[contains(@class, 'j-field-field_') or contains(@class, 'ant-row ant-form-item-row') or contains(@class, 'ui-field-contain')]
    2. 父容器文本 = 标题（无需过滤）
    3. 父容器内找输入框：//input[@type='text' or @type='tel']，无则跳过
    """
    print("\n📌 开始通用表单自动化填写...")
    auto_filled_all = True

    # ---------------------- 1. 按你指定的XPath定位所有父容器 ----------------------
    parent_xpath = "//div[contains(@class, 'j-field-field_') or contains(@class, 'ant-row ant-form-item-row') or contains(@class, 'ui-field-contain')]"

    try:
        # 等待父容器加载并获取所有符合条件的父容器（按页面顺序）
        WebDriverWait(driver, 5).until(EC.presence_of_element_located(('xpath', parent_xpath)))
        parent_containers = driver.find_elements('xpath', parent_xpath)
    except TE:
        print("❌ 未找到任何符合特征的表单父容器")
        return False

    if not parent_containers:
        print("❌ 表单父容器列表为空")
        return False
    print(f"✅ 共识别到 {len(parent_containers)} 个表单字段父容器")

    # ---------------------- 2. 遍历每个父容器，提取标题+匹配输入框 ----------------------
    for group_idx, parent in enumerate(parent_containers, 1):
        print(f"\n=== 处理第 {group_idx} 个字段组 ===")

        # 2.1 提取标题：父容器文本直接作为标题（按你的要求，父容器文本只有标题）
        title_text = parent.text.strip()
        if not title_text:
            print("⚠️ 该父容器无标题文本，跳过")
            auto_filled_all = False
            continue
        print(f"字段标题：{title_text}")

        # 2.2 父容器内找输入框（严格按你的XPath规则）
        try:
            # 仅在当前父容器内定位输入框（相对定位 .//）
            input_ele = parent.find_element(
                'xpath',
                ".//input[@type='text' or @type='tel']"
            )
        except Exception:
            # 无对应输入框，直接跳过
            print("❌ 该字段无文本输入框，跳过")
            auto_filled_all = False
            continue

        # 2.3 双向模糊匹配配置（保留你原有的逻辑）
        matched_key = None
        clean_title = re.sub(r'[:：*★☆\s]+', '', title_text)  # 仅清洗特殊字符，不过滤内容
        for config_key in configs:
            clean_config = re.sub(r'[:：*★☆\s]+', '', config_key.strip())
            if (clean_title in clean_config) or (clean_config in clean_title):
                matched_key = config_key
                break

        # 2.4 填写输入框
        if matched_key and configs[matched_key]:
            try:
                input_ele.clear()  # 清空原有内容
                input_ele.send_keys(configs[matched_key])
                print(f"✅ 已自动填写：{title_text} = {configs[matched_key]}")
            except Exception as e:
                auto_filled_all = False
                print(f"❌ 填写失败：{title_text}（原因：{str(e)[:50]}）")
        else:
            auto_filled_all = False
            print(f"❌ 无匹配配置，需手动填写：{title_text}")

    # ---------------------- 3. 手动提交提示 ----------------------
    print("\n" + "-" * 50)
    if auto_filled_all:
        print("✅ 所有可匹配的文本字段已自动填写完成！")
    else:
        print("⚠️ 部分字段已自动填写，剩余字段需您手动完善")
    print("📢 请您手动检查并点击提交按钮完成表单提交")
    print("-" * 50)

    return auto_filled_all

# ---------------------- 主流程（URL识别+分流执行） ----------------------
if __name__ == '__main__':
    # 1. 获取表单URL
    url = geturl()

    # 2. 检查驱动
    driver_path = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
    if not os.path.exists(driver_path):
        print(f"错误：未找到chromedriver.exe，当前目录：{os.path.dirname(__file__)}")
        exit()

    # 3. 加载配置文件
    try:
        with open('configs.yaml', 'rb') as conf_file:
            configs = yaml.load(conf_file, Loader=yaml.SafeLoader)
    except FileNotFoundError:
        print("错误：未找到configs.yaml")
        exit()
    except yaml.YAMLError as e:
        print(f"配置文件错误：{e}")
        exit()

    # 4. 初始化浏览器
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    try:
        driver = webdriver.Chrome(
            service=Service(driver_path),
            options=options
        )
    except Exception as e:
        print(f"浏览器启动失败：{str(e)}")
        exit()

    # 5. URL识别+分流执行
    try:
        driver.get(url)
        print(f"已打开表单：{url}")

        # 识别表单类型（不区分大小写）- 优化点：简化重复判断，统一执行逻辑
        url_lower = url.lower()
        auto_filled_all = False
        # 合并支持的域名特征，简化判断
        if any(key in url_lower for key in ["jsj.top", "wjx", "jsjform"]):
            auto_filled_all = auto_fill_generic(driver, configs)
        else:
            print("❌ 当前表单链接不支持！目前支持的表单链接包括：")
            print("1.https://xxxx.jsjform.com/x/xxxxxx")
            print("2.https://jsj.top/xx/xxxxxx")
            print("3.https://wjx.cn/xx/xxx")
            print("4.https://wjx.top/xx/xxx")
            print("您可以通过 ctrl+s 保存当前网页结构，后续再增加新的表单自动化逻辑")

        # 统一等待Esc键关闭浏览器
        if auto_filled_all:
            print("完成后按【Esc键】关闭浏览器（按其他键无反应）")
        else:
            print("手动完善未填写字段后，按【Esc键】关闭浏览器")
        keyboard.wait("esc")
        print("已检测到按键，准备关闭浏览器...")
    except Exception as e:
        print(f"出错：{str(e)}")
    finally:
        if driver:
            driver.quit()
            print("浏览器已关闭")
