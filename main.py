import os
import requests
import logging
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv
from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
)

load_dotenv()
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
logging.basicConfig(level=logging.INFO)
USER_TASKS = {}

def find_internal_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        anchor = a.text.strip()
        if not anchor:
            continue
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

def get_status(url):
    try:
        r = requests.get(url, timeout=10, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
        })
        return r.status_code
    except Exception:
        return "ERR"

async def handle_excel(input_path, output_path, context, chat_id):
    wb = load_workbook(input_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    url_idx = None
    for idx, val in enumerate(headers):
        if val and str(val).strip().lower() == 'url':
            url_idx = idx
    if url_idx is None:
        await context.bot.send_message(chat_id=chat_id, text="Không tìm thấy cột 'url' trong file.")
        return
    urls = []
    for row in ws.iter_rows(min_row=2, max_col=url_idx+1):
        url = row[url_idx].value if url_idx < len(row) else None
        if url:
            urls.append(url)

    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.append(["Source URL", "Anchor Text", "Destination URL", "Response Code"])

    for i, src_url in enumerate(urls, 1):
        try:
            resp = requests.get(src_url, timeout=15, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
            })
            if resp.status_code != 200:
                result_ws.append([src_url, "", src_url, resp.status_code])
                await context.bot.send_message(chat_id=chat_id, text=f"{i}/{len(urls)}: {src_url} --> {resp.status_code}")
                continue
            html = resp.text
        except Exception as e:
            result_ws.append([src_url, "", src_url, "ERR"])
            await context.bot.send_message(chat_id=chat_id, text=f"{i}/{len(urls)}: {src_url} --> ERR")
            continue

        base_url = get_base_url(src_url)
        internal_links = find_internal_links(html, base_url)
        if not internal_links:
            result_ws.append([src_url, "", "", "No Internal Link"])
            await context.bot.send_message(chat_id=chat_id, text=f"{i}/{len(urls)}: {src_url} --> No Internal Link")
            continue

        error_count = 0
        for anchor, dst_url in internal_links:
            code = get_status(dst_url)
            result_ws.append([src_url, anchor, dst_url, code])
            if code not in [200, 201, 202]:
                error_count += 1

        await context.bot.send_message(chat_id=chat_id, text=f"{i}/{len(urls)}: {src_url} --> Found {len(internal_links)} internal links, {error_count} lỗi")

    result_wb.save(output_path)
    # DEBUG
    wb_out = load_workbook(output_path)
    ws_out = wb_out.active
    print("==== File output preview ====")
    for r in ws_out.iter_rows(values_only=True):
        print(r)
    print("============================")
    if ws_out.max_row < 2:
        await context.bot.send_message(chat_id=chat_id, text="Không crawl được link nào hợp lệ!")
        return

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Gửi file Excel (.xlsx) gồm cột 'url'.\n"
        "Bot sẽ báo cáo tất cả internal link cùng response code như Screaming Frog.\n"
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

    await update.message.reply_text("Đang xử lý, sẽ báo từng URL...")

    try:
        await handle_excel(input_path, output_path, context, chat_id)
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            await update.message.reply_document(InputFile(output_path, filename="internal_link_report.xlsx"))
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
