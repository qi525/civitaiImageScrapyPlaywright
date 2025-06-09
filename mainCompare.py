# -*- coding: utf-8 -*-
"""
主流程：先注入按钮，确认注入成功后再抓取
"""

import os
import json
import asyncio
import logging
import subprocess
import traceback
from datetime import datetime
from playwright.async_api import async_playwright
# 其它依赖略...

# --- 配置 ---
PROXY = "http://127.0.0.1:10808"
TARGET_URL = "https://civitai.com/images?tags=4"
INJECT_JS_FILE = "控制台注入版.js"
LOG_DIR = "logs"
RESULTS_DIR = "results_civitai"
IMAGE_DIR_BASE = "images_civitai"
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
log_filename = os.path.join(LOG_DIR, f"main_log_{timestamp}.txt")

# --- 日志配置 ---
logger = logging.getLogger("main")
logger.setLevel(logging.INFO)
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
file_handler = logging.FileHandler(log_filename, encoding="utf-8")
console_handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# --- 步骤1：注入按钮 ---
async def inject_scroll_btn(page, js_file=INJECT_JS_FILE, max_retry=10, retry_interval=1):
    if not os.path.exists(js_file):
        logger.error(f"JS 文件不存在: {js_file}")
        return False
    with open(js_file, "r", encoding="utf-8") as f:
        content_js_code = f.read()
    # 多次尝试注入
    for i in range(max_retry):
        logger.info(f"[evaluate] 注入尝试 {i+1}/{max_retry} ...")
        await page.evaluate(content_js_code)
        try:
            await page.wait_for_selector("#ultra-scroll-menu-btn", timeout=1500)
            logger.info("[evaluate] 按钮已成功注入！")
            return True
        except Exception:
            logger.info("[evaluate] 按钮未出现，继续尝试注入...")
            await asyncio.sleep(retry_interval)
    logger.error("[evaluate] 多次注入后仍未检测到按钮，终止流程。")
    return False

# --- 步骤2：爬虫抓取主流程（示例，需补充你的抓取逻辑） ---
async def perform_civitai_scrape(page):
    logger.info("开始执行爬虫抓取逻辑...")
    # 这里写你的抓取逻辑，比如滚动、提取图片等
    await page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
    await asyncio.sleep(2)
    # ...后续抓取逻辑
    logger.info("抓取逻辑结束。")

# --- 主入口 ---
async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            proxy={"server": PROXY} if PROXY else None,
        )
        page = await browser.new_page()
        page.on("console", lambda msg: print("PAGE LOG:", msg.text))
        logger.info(f"打开页面: {TARGET_URL}")
        await page.goto(TARGET_URL)
        await page.wait_for_load_state("load")

        # 先注入按钮
        inject_success = await inject_scroll_btn(page)
        if not inject_success:
            logger.error("按钮注入失败，终止后续抓取。")
            await browser.close()
            return

        # 注入成功后再执行爬虫抓取
        await perform_civitai_scrape(page)

        await browser.close()
        logger.info("浏览器已关闭，流程结束。")

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("用户中断。")
    except Exception as e:
        logger.critical(f"主流程异常: {e}\n{traceback.format_exc()}")
    finally:
        # 自动打开日志文件
        try:
            if os.path.exists(log_filename):
                print(f"打开日志文件: {log_filename}")
                if os.name == 'nt':
                    os.startfile(log_filename)
                elif hasattr(os, "uname") and os.uname().sysname == 'Darwin':
                    subprocess.run(['open', log_filename])
                else:
                    subprocess.run(['xdg-open', log_filename])
        except Exception as e:
            print(f"打开日志文件失败: {e}")