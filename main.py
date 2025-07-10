import os
import requests
import asyncio
import aiohttp
import logging
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from python_dotenv import load_dotenv

from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
)

load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

logging.basicConfig(level=logging.INFO)

# Helper: Detect page type
def detect_page_type(soup):
    if soup.find('div', class_='article-inner'):
        return 'post'
    elif soup.find('div', id='content'):
        return 'page'
    elif soup.find('div', class_='taxonomy-description'):
        return 'category'
    return None

# Helper: Extract content by type
def extract_content(soup, page_type):
    if page_type == 'post':
        node = soup.find('div', class_='article-inner')
    elif page_type == 'page':
        node = soup.find('div', id='content')
    elif page_type == 'category':
        node = soup.find('div', class_='taxonomy-description')
    else:
        return ""
    return node.decode_contents() if node else ""

# Helper: Find internal links in html content
def find_internal_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        anchor = a.text.strip()
        if not anchor:
            continue
        # Only check internal links
        if (href.startswith("/") or base_url in href):
            full_link = href if href.startswith("http") else base_url.rstrip("/") + "/" + href.lstrip("/")
            links.append((anchor, full_link))
    return links

# Async check link status
async def check_link_status(session, url):
    try:
        async with session.get(url, allow_redirects=True, timeout=10) as response:
            return response.status
    except Exception:
        return None

async def process_url(session, url, stt):
    try:
        async with session.get(url, timeout=15) as resp:
            if resp.status != 200:
                return [stt, url, "Toàn bộ URL lỗi", url, "", "", ""]
            html = await resp.text()
    except Exception:
        return [stt, url, "Không truy cập được URL", url, "", "", ""]
    
    soup = BeautifulSoup(html, "html.parser")
    page_type = detect_page_type(soup)
    if not page_type:
        return [stt, url, "Không nhận diện dạng bài", url, "", "", ""]
    content = extract_content(soup, page_type)
    if not content:
        return [stt, url, "Không lấy được nội dung", url, "", "", ""]
    # Lấy base_url
    base_url = url.split("//", 1)[-1].split("/", 1)[0]
    base_url = url.split("//")[0] + "//" + base_url

    # Tìm internal links
    internal_links = find_internal_links(content, base_url)
    result_row = [stt, url]
    anchor_error = 0
    for anchor, link in internal_links:
        code = await check_link_status(session, link)
        if code in [301, 404]:
            anchor_error += 1
            result_row.append(f"{anchor} - {link}")
            if anchor_error == 4:
                break
    # Padding nếu ít hơn 4 lỗi
    while len(result_row) < 6:
        result_row.append("")
    return result_row

async def handle_excel(file_path, output_path):
    wb = load_workbook(file_path)
    ws = wb.active
    urls = [row[0].value for row in ws.iter_rows(min_row=2, min_col=1, max_col=1) if row[0].value]

    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.append(['stt', 'url check', 'anchor bị lỗi và link lỗi 1', 'anchor bị lỗi và link lỗi 2', 'anchor bị lỗi và link lỗi 3', 'anchor bị lỗi và link lỗi 4'])

    async with aiohttp.ClientSession() as session:
        tasks = []
        for i, url in enumerate(urls, 1):
            tasks.append(process_url(session, url, i))
        all_results = await asyncio.gather(*tasks)

        for row in all_results:
            result_ws.append(row)
    result_wb.save(output_path)

# Telegram bot handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Gửi file Excel chứa danh sách URL vào đây.")

async def handle_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.document.mime_type not in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
        await update.message.reply_text("Chỉ nhận file .xlsx.")
        return
    file = await context.bot.get_file(update.message.document.file_id)
    input_path = f"input_{update.message.document.file_id}.xlsx"
    output_path = f"output_{update.message.document.file_id}.xlsx"
    await file.download_to_drive(input_path)

    await update.message.reply_text("Đang xử lý, vui lòng đợi...")

    await handle_excel(input_path, output_path)

    await update.message.reply_document(InputFile(output_path, filename="ketqua_internal_link.xlsx"))
    os.remove(input_path)
    os.remove(output_path)

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()
