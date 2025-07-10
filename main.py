import os
import requests
import logging
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from dotenv import load_dotenv
from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
)

load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
logging.basicConfig(level=logging.INFO)

USER_TASKS = {}

def find_internal_links_from_html(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        anchor = a.text.strip()
        if not anchor:
            continue
        # Logic internal: bắt đầu / hoặc có base_url (domain)
        if href.startswith("/") or base_url in href:
            full_link = href if href.startswith("http") else base_url.rstrip("/") + "/" + href.lstrip("/")
            links.append((anchor, full_link))
    return links

def get_base_url(url):
    try:
        parts = url.split("//", 1)
        main = parts[1].split("/", 1)[0]
        return parts[0] + "//" + main
    except Exception:
        return url

def check_link_status(url):
    try:
        r = requests.get(url, timeout=10, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'})
        return r.status_code
    except Exception:
        return None

def process_url(url, stt):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'}
    try:
        resp = requests.get(url, timeout=15, headers=headers)
        status = resp.status_code
        if status != 200:
            return [stt, url, f"URL lỗi: status {status}", url]
        html = resp.text
    except Exception as e:
        return [stt, url, f"Không truy cập được URL: {str(e)}", url]
    
    base_url = get_base_url(url)
    internal_links = find_internal_links_from_html(html, base_url)
    result_row = [stt, url]
    error_found = False
    for anchor, link in internal_links:
        code = check_link_status(link)
        if code in [301, 404]:
            error_found = True
            result_row.append(f"{anchor} - {link} (status {code})")
    if not error_found:
        result_row.append("OK")
    return result_row

def handle_excel(file_path, output_path):
    wb = load_workbook(file_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    url_idx = None
    for idx, val in enumerate(headers):
        if val and str(val).strip().lower() == 'url':
            url_idx = idx
    if url_idx is None:
        raise Exception("Không tìm thấy cột 'url' trong file.")
    urls = []
    for row in ws.iter_rows(min_row=2, max_col=url_idx+1):
        url = row[url_idx].value if url_idx < len(row) else None
        if url:
            urls.append(url)

    result_rows = []
    max_error = 0

    for i, url in enumerate(urls, 1):
        row = process_url(url, i)
        # trừ 2 cột đầu là stt, url
        num_error = len(row) - 2
        max_error = max(max_error, num_error)
        result_rows.append(row)

    # Header động, số cột lỗi tối đa từng gặp
    header = ['stt', 'url check'] + [f'anchor lỗi {i+1}' for i in range(max_error if max_error > 0 else 1)]
    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.append(header)

    # Fill dòng cho đủ cột
    for row in result_rows:
        while len(row) < len(header):
            row.append("")
        result_ws.append(row)
    result_wb.save(output_path)

    # DEBUG: in toàn bộ nội dung file output trước khi gửi về
    wb_out = load_workbook(output_path)
    ws_out = wb_out.active
    print("==== File output preview ====")
    for r in ws_out.iter_rows(values_only=True):
        print(r)
    print("============================")
    if ws_out.max_row < 2:
        raise Exception("Không crawl được link nào hợp lệ!")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Gửi file Excel (.xlsx) gồm cột 'url'. Không cần cột 'type'.\n"
        "Bot sẽ crawl và kiểm tra toàn bộ internal link cho bạn.\n"
        "Gửi /cancel để dừng tiến trình."
    )

async def handle_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    chat_id = update.effective_chat.id

    USER_TASKS[user_id] = {"cancel": False}

    if update.message.document.mime_type not in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
        await update.message.reply_text("Chỉ nhận file .xlsx.")
        return
    file = await context.bot.get_file(update.message.document.file_id)
    input_path = f"input_{update.message.document.file_id}.xlsx"
    output_path = f"output_{update.message.document.file_id}.xlsx"
    await file.download_to_drive(input_path)

    await update.message.reply_text("Đang xử lý, vui lòng đợi...")

    try:
        handle_excel(input_path, output_path)
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            await update.message.reply_document(InputFile(output_path, filename="ketqua_internal_link.xlsx"))
        else:
            await context.bot.send_message(chat_id=chat_id, text="Không tạo được file kết quả hoặc file kết quả rỗng.")
    except Exception as ex:
        await context.bot.send_message(chat_id=chat_id, text=f"Lỗi: {ex}")
    finally:
        if os.path.exists(input_path): os.remove(input_path)
        if os.path.exists(output_path): os.remove(output_path)
        USER_TASKS.pop(user_id, None)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in USER_TASKS:
        USER_TASKS[user_id]["cancel"] = True
        await update.message.reply_text("Đã nhận lệnh dừng tiến trình.")
    else:
        await update.message.reply_text("Không có tiến trình nào đang chạy.")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()
