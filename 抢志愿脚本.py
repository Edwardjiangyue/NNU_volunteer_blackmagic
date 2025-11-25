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
# 金数据1：https://jsj.top/f/p3RiEU
# 金数据2：企业版问卷，没办法自建
# 问卷星：https://www.wjx.top/vm/YZjvTpC.aspx#

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


# ---------------------- jsj.top自动化逻辑 ----------------------
def auto_fill_jsjtop(driver, configs):
    print("\n📌 识别到金数据链接，开始自动化填写...")
    # 手动定位存放所有项目标题的父节点（你原有的XPath）
    parent_container_xpath = '/html/body/div[1]/div/div/form/div[3]/div'
    parent_container = explicit_find_xpath(driver, parent_container_xpath)
    if not parent_container:
        raise Exception("未找到表单字段容器，请检查父节点XPath")

    index = 2  # 起始索引（你原有的逻辑）
    auto_filled_all = True
    while True:
        # 定位当前索引的字段容器（如div[2]、div[4]）
        field_container_xpath = f'{parent_container_xpath}/div[{index}]'
        field_container = explicit_find_xpath(driver, field_container_xpath)
        if not field_container:
            print("\n所有字段已遍历完成")
            break

        # 提取当前字段的标题（你原有的绝对路径）
        title_xpath = f'{field_container_xpath}/div/div/div[1]/label/span/div'
        title_element = explicit_find_xpath(driver, title_xpath)
        if not title_element:
            print(f"容器div[{index}]未找到标题，跳过")
            auto_filled_all = False
            index += 2
            continue
        title_text = title_element.text.strip()
        if not title_text:
            print(f"容器div[{index}]标题为空，跳过")
            auto_filled_all = False
            index += 2
            continue

        print(f"\n发现字段：{title_text}（容器索引：{index}）")

        # 双向模糊匹配配置（你原有的匹配逻辑）
        matched_key = None
        for config_key in configs:
            if (title_text in config_key) or (config_key in title_text):
                matched_key = config_key
                break

        # 填写逻辑（你原有的输入框定位）
        if matched_key and configs[matched_key]:
            input_xpath = f'{field_container_xpath}//input'
            input_element = explicit_find_xpath(driver, input_xpath)
            if input_element:
                input_element.send_keys(configs[matched_key])
                print(f"✅ 已自动填写：{title_text} = {configs[matched_key]}（配置键：{matched_key}）")
            else:
                auto_filled_all = False
                print(f"❌ 找到配置但未定位到输入框，需手动填写：{title_text}")
        else:
            auto_filled_all = False
            print(f"❌ 无匹配配置，需手动填写：{title_text}")
        index += 2

    # 金数据提交（你原有的提交按钮XPath）
    submit_xpath = '/html/body/div[1]/div/div/form/div[4]/div/button'
    if auto_filled_all:
        print("准备提交...")
        explicit_click_xpath(driver, submit_xpath)
        print("提交成功！")
        time.sleep(3)
    return auto_filled_all


# ---------------------- wjx自动化逻辑（适配问卷星结构） ----------------------
def auto_fill_wjx(driver, configs):
    print("\n📌 识别到问卷星链接，开始自动化填写...")
    # 手动定位存放所有项目标题的父节点（你原有的XPath）
    parent_container_xpath = '/html/body/div[2]/form/div[10]/div[5]/fieldset'
    parent_container = explicit_find_xpath(driver, parent_container_xpath)
    if not parent_container:
        raise Exception("未找到表单字段容器，请检查父节点XPath")

    index = 1  # 起始索引（你原有的逻辑）
    auto_filled_all = True
    while True:
        # 定位当前索引的字段容器（如div[2]、div[4]）
        field_container_xpath = f'{parent_container_xpath}/div[{index}]'
        field_container = explicit_find_xpath(driver, field_container_xpath)
        if not field_container:
            print("\n所有字段已遍历完成")
            break

        # 提取当前字段的标题（你原有的绝对路径）
        title_xpath = f'{field_container_xpath}/div[1]/div[2]'
        title_element = explicit_find_xpath(driver, title_xpath)
        if not title_element:
            print(f"容器div[{index}]未找到标题，跳过")
            auto_filled_all = False
            index += 1
            continue
        title_text = title_element.text.strip()
        if not title_text:
            print(f"容器div[{index}]标题为空，跳过")
            auto_filled_all = False
            index += 1
            continue

        print(f"\n发现字段：{title_text}（容器索引：{index}）")

        # 双向模糊匹配配置（你原有的匹配逻辑）
        matched_key = None
        for config_key in configs:
            if (title_text in config_key) or (config_key in title_text):
                matched_key = config_key
                break

        # 填写逻辑（你原有的输入框定位）
        if matched_key and configs[matched_key]:
            input_xpath = f'{field_container_xpath}//input'
            input_element = explicit_find_xpath(driver, input_xpath)
            if input_element:
                input_element.send_keys(configs[matched_key])
                print(f"✅ 已自动填写：{title_text} = {configs[matched_key]}（配置键：{matched_key}）")
            else:
                auto_filled_all = False
                print(f"❌ 找到配置但未定位到输入框，需手动填写：{title_text}")
        else:
            auto_filled_all = False
            print(f"❌ 无匹配配置，需手动填写：{title_text}")
        index += 1

    # 问卷星隐私条款勾选（标准结构）
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.5)
        # 优先点击label（兼容问卷星标准结构）
        privacy_xpath = '//label[@for="checkxiexi"]'
        privacy = explicit_find_xpath(driver, privacy_xpath)
        if privacy:
            explicit_click_xpath(driver, privacy_xpath)
            print("✅ 已勾选隐私条款")
    except Exception as e:
            print(f"❌ 勾选隐私条款失败：{str(e)}")
            auto_filled_all = False


    # 问卷星提交按钮（通用定位）
    submit_xpath = '//div[@id="ctlNext" and contains(@class, "submitbtn")]'
    if auto_filled_all:
        print("准备提交...")
        explicit_click_xpath(driver, submit_xpath)
        print("提交成功！")
        time.sleep(3)
    return auto_filled_all


# ---------------------- jsjform自动化逻辑 ----------------------

def auto_fill_jsjform(driver, configs):
    print("\n📌 识别到金数据链接，开始自动化填写...")
    # 手动定位存放所有项目标题的父节点（你原有的XPath）
    parent_container_xpath = '/html/body/div[2]/div[2]/div[1]/form/div/div[2]'
    parent_container = explicit_find_xpath(driver, parent_container_xpath)
    if not parent_container:
        raise Exception("未找到表单字段容器，请检查父节点XPath")

    index = 1  # 起始索引（你原有的逻辑）
    auto_filled_all = True
    while True:
        # 定位当前索引的字段容器（如div[2]、div[4]）
        field_container_xpath = f'{parent_container_xpath}/div[{index}]'
        field_container = explicit_find_xpath(driver, field_container_xpath)
        if not field_container:
            print("\n所有字段已遍历完成")
            break

        # 提取当前字段的标题（你原有的绝对路径）
        title_xpath = f'{field_container_xpath}/div[1]'
        title_element = explicit_find_xpath(driver, title_xpath)
        if not title_element:
            print(f"容器div[{index}]未找到标题，跳过")
            auto_filled_all = False
            index += 1
            continue
        title_text = title_element.text.strip()
        if not title_text:
            print(f"容器div[{index}]标题为空，跳过")
            auto_filled_all = False
            index += 1
            continue

        print(f"\n发现字段：{title_text}（容器索引：{index}）")

        # 双向模糊匹配配置（你原有的匹配逻辑）
        matched_key = None
        for config_key in configs:
            if (title_text in config_key) or (config_key in title_text):
                matched_key = config_key
                break

        # 填写逻辑（你原有的输入框定位）
        if matched_key and configs[matched_key]:
            input_xpath = f'{field_container_xpath}//input'
            input_element = explicit_find_xpath(driver, input_xpath)
            if input_element:
                input_element.send_keys(configs[matched_key])
                print(f"✅ 已自动填写：{title_text} = {configs[matched_key]}（配置键：{matched_key}）")
            else:
                auto_filled_all = False
                print(f"❌ 找到配置但未定位到输入框，需手动填写：{title_text}")
        else:
            auto_filled_all = False
            print(f"❌ 无匹配配置，需手动填写：{title_text}")
        index += 1

    # 金数据提交按钮（通用定位）
    submit_xpath = '//button[@type="submit" and @data-type="primary" and .//span[text()="提交"]]'
    if auto_filled_all:
        print("准备提交...")
        explicit_click_xpath(driver, submit_xpath)
        print("提交成功！")
        time.sleep(3)
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

        # 识别表单类型（不区分大小写）
        if "jsj.top" in url.lower():
            auto_filled_all = auto_fill_jsjtop(driver, configs)
        elif "wjx" in url.lower():
            auto_filled_all = auto_fill_wjx(driver, configs)
        elif "jsjform" in url.lower():
            auto_filled_all = auto_fill_jsjform(driver, configs)
        else:
            print("❌ 当前表单链接不支持！目前支持的表单链接包括：")
            print("1.https://xxxx.jsjform.com/x/xxxxxx")
            print("2.https://jsj.top/xx/xxxxxx")
            print("3.https://wjx.cn/xx/xxx")
            print("4.https://wjx.top/xx/xxx")
            print("您可以通过 ctrl+s 保存当前网页结构，后续再增加新的表单自动化逻辑")
            auto_filled_all = False

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
