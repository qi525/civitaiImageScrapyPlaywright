# -*- coding: utf-8 -*-
"""
Civitai图片爬虫主流程：只用核心滚动代码，抓取所有图片信息并保存到Excel
"""

import os
import json
import asyncio
import aiohttp
from datetime import datetime
import traceback
import logging
import subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import hashlib
import aiofiles
import re
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
import playwright._impl._errors
import time

# --- 配置 ---
PROXY = "http://127.0.0.1:10808"
TARGET_URL = "https://civitai.com/images?tags=4"
LOG_DIR = "logs"
RESULTS_DIR = "results_civitai"
IMAGE_DIR_BASE = "images_civitai"
KEYWORD_TARGET_FILE = "urlTarget.txt"
DOWNLOAD_HISTORY_FILE = "download_history_civitai.json" # Original MD5-based history
HISTORY_IMG_URL_FILE = "history_img_url_history.xlsx" # New URL/Path based history

timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
log_filename = os.path.join(LOG_DIR, f"civitai_scraper_log_{timestamp}.txt")
excel_filename = os.path.join(RESULTS_DIR, f"civitai_image_results_{timestamp}.xlsx")

# --- 日志配置 ---
logger = logging.getLogger('civitai_scraper')
logger.setLevel(logging.INFO)
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)
if not os.path.exists(IMAGE_DIR_BASE):
    os.makedirs(IMAGE_DIR_BASE)
file_handler = logging.FileHandler(log_filename, encoding='utf-8')
console_handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# --- 工具函数 ---
def calculate_md5(data_bytes):
    return hashlib.md5(data_bytes).hexdigest()

def calculate_url_md5(url):
    """Calculates MD5 for a given URL string, used for consistent naming."""
    return hashlib.md5(url.encode('utf-8')).hexdigest()

# Global history for MD5-based downloaded content
download_history = {}
# Global history for URL-based downloaded content (new)
url_download_history = {}

def load_download_history(filepath):
    global download_history
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                download_history = json.load(f)
            logger.info(f"Loaded download history from {filepath}")
        except Exception as e:
            logger.warning(f"Failed to load download history (MD5): {e}")
            download_history = {}
    else:
        logger.info(f"Download history file '{filepath}' not found. Starting with empty history.")
        download_history = {}

def save_download_history(filepath):
    global download_history
    try:
        dir_name = os.path.dirname(filepath)
        if dir_name:
            os.makedirs(dir_name, exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(download_history, f, indent=4)
        logger.info(f"Saved download history to {filepath}")
    except Exception as e:
        logger.error(f"Error saving download history (MD5) to '{filepath}': {e}\n{traceback.format_exc()}")

def load_url_history(filepath):
    global url_download_history
    if os.path.exists(filepath):
        try:
            wb = load_workbook(filepath)
            ws = wb.active
            header = [cell.value for cell in ws[1]]
            for row_idx in range(2, ws.max_row + 1):
                row_data = {header[col_idx]: ws.cell(row=row_idx, column=col_idx+1).value for col_idx in range(len(header))}
                thumb_url = row_data.get("Thumbnail URL")
                orig_page_url = row_data.get("Original Page URL")
                local_path = row_data.get("Local Image Path")
                if thumb_url and orig_page_url and local_path:
                    # Using a combined key for uniqueness
                    key = f"{thumb_url}|{orig_page_url}"
                    url_download_history[key] = {
                        "local_path": local_path,
                        "image_md5": row_data.get("MD5 (Content)") # This might be useful for verification if needed
                    }
            logger.info(f"Loaded URL download history from {filepath}")
        except Exception as e:
            logger.warning(f"Failed to load URL download history from {filepath}: {e}")
            url_download_history = {}
    else:
        logger.info(f"URL download history file '{filepath}' not found. Starting with empty history.")
        url_download_history = {}

def save_url_history(filepath):
    global url_download_history
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Image Download History"
        headers = ["Thumbnail URL", "Original Page URL", "Local Image Path", "MD5 (Content)"]
        ws.append(headers)

        for key, data in url_download_history.items():
            thumb_url, orig_page_url = key.split('|', 1) # Split only on the first '|'
            ws.append([thumb_url, orig_page_url, data["local_path"], data.get("image_md5", "")])

        for col_idx, header in enumerate(headers):
            max_length = len(header)
            column_letter = get_column_letter(col_idx + 1)
            for r_idx in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=r_idx, column=col_idx + 1).value
                if cell_value:
                    cell_len = len(str(cell_value))
                    if cell_len > max_length:
                        max_length = cell_len
            adjusted_width = (max_length + 2) * 1.2
            if adjusted_width > 100:
                adjusted_width = 100
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(filepath)
        logger.info(f"Saved URL download history to {filepath}")
    except Exception as e:
        logger.error(f"Error saving URL download history to '{filepath}': {e}\n{traceback.format_exc()}")

def read_urls_from_file(filepath):
    urls = []
    if not os.path.exists(filepath):
        logger.error(f"Error: Target URL file '{filepath}' not found. 请创建并添加URL。")
        return []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                url = line.strip()
                if url and url.startswith("http"):
                    urls.append(url)
        if not urls:
            logger.warning(f"Warning: Target URL file '{filepath}' is empty or contains no valid URLs.")
        return urls
    except Exception as e:
        logger.error(f"Error reading URLs from '{filepath}': {e}\n{traceback.format_exc()}")
        return []

async def process_image_data(image_url, original_page_url, base_folder_path):
    local_filename = None
    image_content_md5 = None
    image_bytes = None

    if not image_url:
        logger.warning("Empty image URL skipped.")
        return None, None

    # Determine file extension
    url_without_query = image_url.split('?')[0]
    file_extension = url_without_query.split('.')[-1].lower()
    if not file_extension or len(file_extension) > 5 or not file_extension.isalpha():
        file_extension = 'jpg'

    # Create a unique key for the URL history
    unique_url_key = f"{image_url}|{original_page_url}"

    # --- Step 1: Check URL/Path history first to avoid re-downloading ---
    if unique_url_key in url_download_history:
        existing_info = url_download_history[unique_url_key]
        existing_path = existing_info["local_path"]
        
        # Infer the expected filename based on how we name files (MD5 of content, if available)
        # If we have an MD5 from history, we can assume the filename.
        expected_filename_from_history_md5 = None
        if existing_info.get("image_md5"):
            expected_filename_from_history_md5 = os.path.join(base_folder_path, f"{existing_info['image_md5']}.{file_extension}")
        else:
            # If MD5 not in history, but we have a path, try to derive from path
            # This is a fallback if the MD5 was missing in the history record itself.
            expected_filename_from_history_md5 = existing_path # In this case, we just check if the path itself is valid

        if os.path.exists(existing_path) and (not expected_filename_from_history_md5 or os.path.basename(existing_path) == os.path.basename(expected_filename_from_history_md5)):
            # Log the skip with the relevant info
            logger.info(f"Image '{image_url}' found in history with local path '{existing_path}'. File exists and name matches. Skipping download and content MD5 calculation.")
            return existing_path, existing_info.get("image_md5") # Return the MD5 from history

    # --- Step 2: If not in URL history or file missing, proceed to download and calculate MD5 ---
    if image_url.startswith('http'):
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(image_url, proxy=PROXY if PROXY else None, timeout=30.0) as response:
                    response.raise_for_status()
                    image_bytes = await response.read()
                    image_content_md5 = calculate_md5(image_bytes) # Calculate MD5 only for new downloads

                    # Check the original MD5-based download history (secondary check for content de-duplication)
                    if image_content_md5 in download_history:
                        existing_path = download_history[image_content_md5]
                        if os.path.exists(existing_path):
                            logger.info(f"Downloaded image content (MD5: {image_content_md5}) already exists at: {existing_path}. Skipping save and using existing path.")
                            # Update URL history with this path, even if found via MD5 history
                            url_download_history[unique_url_key] = {"local_path": existing_path, "image_md5": image_content_md5}
                            return existing_path, image_content_md5

                    local_filename = os.path.join(base_folder_path, f"{image_content_md5}.{file_extension}")
                    os.makedirs(os.path.dirname(local_filename), exist_ok=True)
                    async with aiofiles.open(local_filename, 'wb') as f:
                        await f.write(image_bytes)
                    logger.info(f"Image downloaded and saved: {local_filename}")

                    # Update both histories
                    download_history[image_content_md5] = local_filename
                    url_download_history[unique_url_key] = {"local_path": local_filename, "image_md5": image_content_md5}
                    return local_filename, image_content_md5
        except Exception as e:
            logger.error(f"Error downloading image {image_url}: {e}")
    else:
        logger.warning(f"Unsupported image URL format (not http/https): {image_url[:100]}...")
    return None, None

# Helper to parse counts with 'k'
def parse_count_with_k(count_str):
    if not count_str:
        return 0
    count_str = count_str.strip().lower()
    if 'k' in count_str:
        try:
            return int(float(count_str.replace('k', '')) * 1000)
        except ValueError:
            return 0
    else:
        try:
            # Use regex to find only digits (and possibly a dot for k)
            match = re.search(r'[\d.]+', count_str)
            if match:
                return int(float(match.group(0)))
            return 0
        except ValueError:
            return 0

# --- 核心爬虫流程 ---
async def performCivitaiImageScrape(context, target_url):
    async_name = asyncio.current_task().get_name()
    # --- Cookie 注入 ---
    try:
        if os.path.exists("cookies.json"):
            with open("cookies.json", "r", encoding="utf-8") as f:
                cookies = json.load(f)
                for cookie in cookies:
                    cookie_same_site = {'strict': 'Strict', 'Lax': 'Lax', 'none': 'None'}.get(cookie.get('sameSite'), None)
                    if cookie_same_site in ['Strict', 'Lax', 'None']:
                        cookie['sameSite'] = cookie_same_site
                    else:
                        if 'sameSite' in cookie:
                            del cookie['sameSite']
                await context.add_cookies(cookies)
            logger.info(f"{async_name} -> Cookies loaded and added to context.")
        else:
            logger.warning(f"{async_name} -> Warning: cookies.json not found. Proceeding without cookies.")
    except Exception as e:
        logger.error(f"{async_name} -> Error loading cookies: {e}")

    page = await context.new_page()
    try:
        logger.info(f"{async_name} -> Navigating to {target_url}")
        await page.goto(target_url, timeout=60000, wait_until="domcontentloaded")
        logger.info(f"{async_name} -> Successfully navigated to {target_url}")
    except Exception as e:
        logger.error(f"{async_name} -> Error navigating to {target_url}: {e}")
        await page.close()
        return

    civitai_image_folder_path = os.path.join(IMAGE_DIR_BASE, "downloaded_images")
    if not os.path.exists(civitai_image_folder_path):
        os.makedirs(civitai_image_folder_path)
        logger.info(f"{async_name} -> Created base image folder for Civitai: {civitai_image_folder_path}")

    scrape_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    processed_image_detail_urls = set()  # 放在 while 循环外
    scroll_attempts = 0
    max_scroll_attempts = 80  # 原来是30，改为200，加大滚动次数
    no_new_images_count = 0
    no_new_images_start_time = None  # 新增：无新图片开始的时间

    image_boxes_selector = '#main > div > div > div > main > div.z-10.m-0.flex-1.mantine-1avyp1d > div > div > div > div:nth-child(2) > div.mx-auto.flex.justify-center.gap-4'
    keyword_input_selector = 'header input'  # 关键词输入框选择器
    current_keyword = "N/A"
    try:
        keyword_input_element = page.locator(keyword_input_selector)
        if await keyword_input_element.is_visible():
            input_value = await keyword_input_element.get_attribute('value')
            if input_value:
                current_keyword = input_value
                logger.info(f"{async_name} -> Found keyword in input field: '{current_keyword}'")
    except Exception as e:
        logger.warning(f"{async_name} -> Could not find or extract keyword: {e}")

    while scroll_attempts < max_scroll_attempts:
        scroll_attempts += 1
        logger.info(f"{async_name} -> Scroll attempt {scroll_attempts}/{max_scroll_attempts}...")
        await page.evaluate("""
            document.querySelectorAll('*').forEach(function(el) {
                if (el.scrollHeight > el.clientHeight) el.scrollTop += 40;
            });
        """)
        #await asyncio.sleep(0.05)

        # 只抓取目标div下的img
        page_html = await page.content()
        soup = BeautifulSoup(page_html, "html.parser")
        target_div = soup.select_one("div.mx-auto.flex.justify-center.gap-4")
        if not target_div:
            logger.warning("目标div未找到，跳过本次循环")
            continue
        
        # Find all individual image card containers within the target_div
        # Based on big big box.html and box.html, the main card container seems to be:
        # <div class="relative flex overflow-hidden rounded-md ... flex-col border" id="...">
        # or similar with class flex-col border and relative flex overflow-hidden
        image_card_containers = target_div.find_all("div", class_=lambda x: x and "flex-col border" in x and "relative flex overflow-hidden" in x)

        newly_processed_this_scroll = 0
        for card_container in image_card_containers:
            thumbnail_url = ""
            original_page_url = ""
            
            # Extract thumbnail URL and original page URL
            img_element = card_container.find("img", class_="EdgeImage_image__iH4_q")
            if img_element:
                thumbnail_url = img_element.get("src")
                parent_a = img_element.find_parent("a")
                if parent_a:
                    original_page_url = parent_a.get("href")
                    if original_page_url and not original_page_url.startswith("http"):
                        original_page_url = f"https://civitai.com{original_page_url}"

            if not thumbnail_url or not thumbnail_url.startswith("http"):
                continue

            # De-duplication check
            unique_key_for_scrape_tracking = thumbnail_url + "|" + original_page_url
            if unique_key_for_scrape_tracking in processed_image_detail_urls:
                continue
            processed_image_detail_urls.add(unique_key_for_scrape_tracking)
            newly_processed_this_scroll += 1

            # --- Extract Button Counts ---
            like_count = 0
            heart_count = 0
            laugh_count = 0
            sad_count = 0
            tip_count = 0

            # Find the div containing the buttons
            # This seems to be <div class="flex items-center justify-center gap-1 justify-between p-2">
            buttons_container = card_container.find("div", class_=lambda x: x and "flex items-center justify-center" in x and "p-2" in x and "gap-1" in x)
            
            if buttons_container:
                # Find all buttons within this container
                buttons = buttons_container.find_all("button", class_=lambda x: x and ("mantine-UnstyledButton-root" in x or "mantine-Button-root" in x))
                for button in buttons:
                    # Handle standard emoji buttons (like, heart, laugh, sad)
                    label_span = button.find("span", class_="mantine-Button-label")
                    if label_span:
                        emoji_div = label_span.find("div", class_="mantine-Text-root")
                        
                        # Get all text from the span and then try to extract the number
                        # This covers cases where the number is directly in the span, not in a separate div
                        full_label_text = label_span.get_text(separator=' ', strip=True)
                        
                        # Use regex to find digits in the full_label_text
                        match = re.search(r'(\d[\d\.]*[kK]?)', full_label_text)
                        
                        count = 0
                        if match:
                            count_str = match.group(1)
                            count = parse_count_with_k(count_str)

                        if emoji_div: # Ensure emoji div exists
                            if "👍" in emoji_div.text:
                                like_count = count
                            elif "❤️" in emoji_div.text:
                                heart_count = count
                            elif "😂" in emoji_div.text:
                                laugh_count = count
                            elif "😢" in emoji_div.text:
                                sad_count = count
                    
                    # Special handling for the tip button (it's a badge, not a simple button label)
                    # This part remains unchanged as per your instruction
                    tip_badge = button.find("div", class_=lambda x: x and "mantine-Badge-root" in x)
                    if tip_badge:
                        # Check for the lightning bolt SVG
                        if tip_badge.find("svg", class_=lambda x: x and "tabler-icon-bolt" in x):
                            tip_text_div = tip_badge.find("div", class_="mantine-Text-root") # The count is in this div
                            if tip_text_div:
                                tip_str = tip_text_div.text.strip()
                                tip_count = parse_count_with_k(tip_str)

            # Download image (or skip if in history)
            local_image_path, image_md5 = await process_image_data(thumbnail_url, original_page_url, civitai_image_folder_path)
            if local_image_path:
                abs_path = os.path.abspath(local_image_path)
                if os.name == 'nt':
                    abs_path = abs_path.replace('\\', '/')
                    local_image_hyperlink = f"file:///{abs_path}"
                else:
                    local_image_hyperlink = f"file://{abs_path}"
            else:
                local_image_hyperlink = ""
            
            result_data = {
                "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "搜索URL": target_url,
                "缩略图URL": thumbnail_url,
                "本地缩略图路径": os.path.abspath(local_image_path) if local_image_path else "",
                "本地缩略图超链接": local_image_hyperlink,
                "原始图片详情页链接": original_page_url,
                "点赞数": like_count,
                "爱心数": heart_count,
                "笑哭数": laugh_count,
                "伤心数": sad_count,
                "打赏数": tip_count,
                "关键词": current_keyword
            }
            async with data_lock:
                all_search_results_data.append(result_data)

        if newly_processed_this_scroll == 0 and len(image_card_containers) > 0: # Check if any cards were found at all
            if no_new_images_start_time is None:
                no_new_images_start_time = time.time()
            elapsed = time.time() - no_new_images_start_time
            logger.info(f"{async_name} -> No new images processed this scroll. Consecutive no new images: {no_new_images_count}, elapsed: {elapsed:.1f}s")
            if elapsed >= 20:
                logger.info(f"{async_name} -> No new images for 20 seconds. Stopping scrolling.")
                break
        else:
            no_new_images_start_time = None  # 有新图片就重置
    await page.close()
    logger.info(f"{async_name} -> Page closed for {target_url}.")

# --- 主入口 ---
all_search_results_data = []
data_lock = asyncio.Lock()

async def main():
    load_download_history(DOWNLOAD_HISTORY_FILE)
    load_url_history(HISTORY_IMG_URL_FILE) # Load the new URL-based history
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            proxy={"server": PROXY} if PROXY else None,
            timeout=60000
        )
        context = await browser.new_context(
            viewport={'width': 2560, 'height': 1440}
        )
        target_urls = read_urls_from_file(KEYWORD_TARGET_FILE)
        if not target_urls:
            logger.error(f"No valid URLs found in {KEYWORD_TARGET_FILE}. Please add URLs to scrape.")
            await browser.close()
            return
        # 直接抓取所有URL
        tasks = [performCivitaiImageScrape(context, url) for url in target_urls]
        await asyncio.gather(*tasks)
        await browser.close()
        logger.info("Browser closed. Script finished scraping data.")

    # --- Excel 导出 ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Civitai图片结果"
    headers = ["抓取时间", "搜索URL", "缩略图URL", "本地缩略图路径", "本地缩略图超链接", "原始图片详情页链接", "点赞数", "爱心数", "笑哭数", "伤心数", "打赏数", "关键词"]
    ws.append(headers)
    hyperlink_font = Font(color="0000FF", underline="single")
    for row_data in all_search_results_data:
        row = []
        for header in headers:
            if header == "本地缩略图超链接":
                row.append("点击打开缩略图")
            else:
                row.append(row_data.get(header, ""))
        ws.append(row)
        current_row_idx = ws.max_row
        search_url = row_data.get("搜索URL")
        if search_url:
            cell_search_url = ws.cell(row=current_row_idx, column=headers.index("搜索URL") + 1)
            cell_search_url.value = search_url
            cell_search_url.hyperlink = search_url
            cell_search_url.font = hyperlink_font
        thumbnail_url = row_data.get("缩略图URL")
        if thumbnail_url:
            cell_thumbnail_url = ws.cell(row=current_row_idx, column=headers.index("缩略图URL") + 1)
            cell_thumbnail_url.value = thumbnail_url
            cell_thumbnail_url.hyperlink = thumbnail_url
            cell_thumbnail_url.font = hyperlink_font
        local_image_hyperlink_url = row_data.get("本地缩略图超链接")
        if local_image_hyperlink_url:
            cell_local_image_hyperlink = ws.cell(row=current_row_idx, column=headers.index("本地缩略图超链接") + 1)
            cell_local_image_hyperlink.value = "点击打开缩略图"
            cell_local_image_hyperlink.hyperlink = Hyperlink(ref=local_image_hyperlink_url)
            cell_local_image_hyperlink.font = hyperlink_font
        original_page_link = row_data.get("原始图片详情页链接")
        if original_page_link:
            cell_original_page_link = ws.cell(row=current_row_idx, column=headers.index("原始图片详情页链接") + 1)
            cell_original_page_link.value = original_page_link
            cell_original_page_link.hyperlink = original_page_link
            cell_original_page_link.font = hyperlink_font
    for col_idx, header in enumerate(headers):
        max_length = len(header)
        column_letter = get_column_letter(col_idx + 1)
        for r_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=r_idx, column=col_idx + 1).value
            if cell_value:
                cell_len = len(str(cell_value))
                if cell_len > max_length:
                    max_length = cell_len
            adjusted_width = (max_length + 2) * 1.2
            if adjusted_width > 100:
                adjusted_width = 100
            ws.column_dimensions[column_letter].width = adjusted_width
    wb.save(excel_filename)
    logger.info(f"Results saved to Excel: {excel_filename}")
    save_download_history(DOWNLOAD_HISTORY_FILE)
    save_url_history(HISTORY_IMG_URL_FILE) # Save the new URL-based history

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Script interrupted by user.")
    except Exception as e:
        logger.critical(f"An unhandled error occurred in main: {e}\n{traceback.format_exc()}")
    finally:
        for handler in logger.handlers[:]:
            try:
                handler.flush()
                handler.close()
                logger.removeHandler(handler)
            except Exception as e:
                print(f"Error closing log handler: {e}")
        try:
            if os.path.exists(log_filename):
                print(f"Attempting to open log file: {log_filename}")
                if os.name == 'nt':
                    os.startfile(log_filename)
                elif hasattr(os, "uname") and os.uname().sysname == 'Darwin':
                    subprocess.run(['open', log_filename])
                else:
                    subprocess.run(['xdg-open', log_filename])
            else:
                print(f"Log file not found: {log_filename}")
        except Exception as e:
            print(f"Error opening log file {log_filename}: {e}")
        try:
            if os.path.exists(excel_filename):
                print(f"Attempting to open Excel file: {excel_filename}")
                if os.name == 'nt':
                    os.startfile(excel_filename)
                elif hasattr(os, "uname") and os.uname().sysname == 'Darwin':
                    subprocess.run(['open', excel_filename])
                else:
                    subprocess.run(['xdg-open', excel_filename])
            else:
                print(f"Excel file not found: {excel_filename}")
        except Exception as e:
            print(f"Error opening Excel file {excel_filename}: {e}")
        try:
            if os.path.exists(HISTORY_IMG_URL_FILE):
                print(f"Attempting to open URL history Excel file: {HISTORY_IMG_URL_FILE}")
                if os.name == 'nt':
                    os.startfile(HISTORY_IMG_URL_FILE)
                elif hasattr(os, "uname") and os.uname().sysname == 'Darwin':
                    subprocess.run(['open', HISTORY_IMG_URL_FILE])
                else:
                    subprocess.run(['xdg-open', HISTORY_IMG_URL_FILE])
            else:
                print(f"URL history Excel file not found: {HISTORY_IMG_URL_FILE}")
        except Exception as e:
            print(f"Error opening URL history Excel file {HISTORY_IMG_URL_FILE}: {e}")