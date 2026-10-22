# -*- coding: utf-8 -*-
"""
استخراج خودکار سوالات از PDF با استفاده از Gemini (نسخه بازنویسی‌شده)
======================================================================
تغییرات اصلی نسبت به نسخه قبلی:
  1. کتابخانه قدیمی و منسوخ‌شده google-generativeai حذف و با کتابخانه
     جدید و رسمی google-genai جایگزین شد (علت اصلی خطای 403 روی
     آدرس discovery همین کتابخانه قدیمی بود).
  2. مدل gemini-1.5-flash (که کاملاً از رده خارج شده و هر درخواستی به
     آن خطای 404 می‌دهد) با gemini-3.1-flash-lite (نسخه فعلی و پایدار
     flash-lite) جایگزین شد.
  3. کلید API دیگر داخل کد نوشته نمی‌شود؛ از فایل .env یا متغیر محیطی
     خوانده می‌شود (به بخش امنیتی در پیام همراه این فایل مراجعه کنید).
  4. هر صفحه پی‌دی‌اف مستقیماً و به‌صورت inline به مدل ارسال می‌شود
     (بدون آپلود/پولینگ روی سرور گوگل)، مگر این‌که فایل بزرگ باشد.
  5. رابط کاربری گرافیکی کامل با نوار پیشرفت و لاگ زنده.
  6. قابلیت ادامه (Resume): اگر برنامه وسط کار متوقف شود، در اجرای
     بعدی فقط صفحات باقی‌مانده پردازش می‌شوند.
  7. ذخیره خودکار فایل Word پس از هر صفحه (از دست رفتن اطلاعات در صورت
     قطع برنامه را از بین می‌برد).
  8. دسته‌بندی هوشمند خطاها (خطای کلید/مدل بلافاصله متوقف می‌شود، خطای
     محدودیت نرخ با backoff نمایی دوباره امتحان می‌شود).
"""

import os
import re
import sys
import time
import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from pypdf import PdfReader, PdfWriter
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from dotenv import load_dotenv, set_key

from google import genai
from google.genai import types
from google.genai import errors as genai_errors


# ==============================================================================
# تنظیمات قابل ویرایش
# ==============================================================================

# اگر گوگل بعداً نام مدل flash-lite را عوض کرد، فقط همین خط را تغییر بده.
# گزینه‌های دیگر در همان کمبوباکس داخل برنامه هم قابل انتخاب/ویرایش هستند.
MODEL_NAME_DEFAULT = "gemini-3.5-flash-lite"

# اگر این فونت روی سیستم نصب نیست، به‌جایش مثلاً "Tahoma" یا "Calibri" بگذار.
PERSIAN_FONT = "B Nazanin"

MAX_RETRIES = 4
BASE_BACKOFF_SECONDS = 5
REQUEST_DELAY_SECONDS = 3          # فاصله امن بین درخواست‌های پیاپی
INLINE_SIZE_LIMIT_BYTES = 15 * 1024 * 1024   # زیر این حجم مستقیم ارسال می‌شود

SCRIPT_DIR = Path(__file__).resolve().parent
ENV_PATH = SCRIPT_DIR / ".env"

EXTRACTION_PROMPT = """تو یک دستیار متخصص در پردازش اسناد آموزشی و استخراج متن فارسی هستی.
وظیفه تو استخراج دقیق و کامل تمام سوالات موجود در این صفحه (چهارگزینه‌ای یا هر نوع دیگر) است.

قوانین:
۱. ساختار سوالات، شماره‌گذاری، و چینش گزینه‌ها (الف/ب/پ/ت یا ۱/۲/۳/۴ یا A/B/C/D، هرکدام که در سند آمده) را دقیقاً همان‌طور که هست حفظ کن.
۲. متن را عیناً از روی سند بازنویسی کن؛ چیزی را ترجمه، خلاصه، تفسیر یا تصحیح نکن.
۳. اگر سوالی از صفحه قبل یا بعد ادامه دارد، فقط همان بخشی را که در همین صفحه است بنویس.
۴. اگر هیچ سوالی در این صفحه نیست (مثلاً فقط تصویر، متن توضیحی یا صفحه خالی است)، فقط دقیقاً همین عبارت را بنویس: [هیچ سوالی در این صفحه یافت نشد]
۵. هیچ‌گونه متن اضافه، سلام، توضیح، عنوان، بررسی یا نتیجه‌گیری اضافه نکن.
۶. خروجی را فقط و فقط داخل یک بلوک کد متنی (```) قرار بده، بدون هیچ متنی قبل یا بعد از آن."""

SUCCESS_HEADER_RE = re.compile(r'^--- سوالات صفحه (\d+) ---$')

FATAL_STOP_CODES = {401, 403, 404}   # خطاهایی که تکرارشان بی‌فایده است


# ==============================================================================
# توابع کمکی — Word (RTL) و PDF و پارس ورودی
# ==============================================================================

def set_rtl_paragraph(paragraph):
    """راست‌چین و راست‌به‌چپ کردن یک پاراگراف Word برای متن فارسی."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    pPr.append(OxmlElement('w:bidi'))


def apply_persian_font(run, size=12, bold=False):
    """اعمال فونت فارسی هم برای حروف معمولی و هم برای حروف پیچیده (راست‌به‌چپ)."""
    run.font.rtl = True
    run.font.size = Pt(size)
    run.font.name = PERSIAN_FONT
    run.bold = bold
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:cs'), PERSIAN_FONT)
    rFonts.set(qn('w:ascii'), PERSIAN_FONT)
    rFonts.set(qn('w:hAnsi'), PERSIAN_FONT)


def parse_pages_input(page_input_str):
    """پارس امن یک رشته صفحات ناپیوسته مثل '250-255, 260, 270-275' به لیست مرتب."""
    pages = []
    for part in page_input_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            start_s, end_s = part.split('-')
            start, end = int(start_s.strip()), int(end_s.strip())
            if start > end:
                start, end = end, start
            pages.extend(range(start, end + 1))
        else:
            pages.append(int(part))
    if not pages:
        raise ValueError("هیچ صفحه‌ای در ورودی پیدا نشد")
    return sorted(set(pages))


def is_valid_name_component(name):
    """بررسی این‌که نام پروژه برای نام‌گذاری فایل/پوشه در ویندوز مجاز است."""
    forbidden = set('<>:"/\\|?*')
    return len(name) > 0 and not any(ch in forbidden for ch in name)


def remove_stale_error_paragraph(doc, page_num):
    """اگر برای این صفحه قبلاً یک پاراگراف خطا ثبت شده (از اجرای ناقص قبلی)، پاکش کن."""
    target_text = f"--- خطا در استخراج سوالات صفحه {page_num} ---"
    for p in doc.paragraphs:
        if p.text.strip() == target_text:
            p._element.getparent().remove(p._element)
            break


def scan_completed_pages(word_path):
    """برگرداندن شماره صفحاتی که قبلاً با موفقیت استخراج شده‌اند (برای ادامه کار)."""
    completed = set()
    if not os.path.exists(word_path):
        return completed
    try:
        doc = Document(word_path)
    except Exception:
        return completed
    for p in doc.paragraphs:
        m = SUCCESS_HEADER_RE.match(p.text.strip())
        if m:
            completed.add(int(m.group(1)))
    return completed


# ==============================================================================
# ارتباط با Gemini
# ==============================================================================

def call_gemini_for_page(client, model_name, pdf_path):
    """ارسال یک PDF تک‌صفحه‌ای به Gemini و برگرداندن متن استخراج‌شده (تمیزشده)."""
    uploaded_name = None
    try:
        size = os.path.getsize(pdf_path)
        if size <= INLINE_SIZE_LIMIT_BYTES:
            with open(pdf_path, "rb") as f:
                data = f.read()
            contents = [types.Part.from_bytes(data=data, mime_type="application/pdf")]
        else:
            uploaded = client.files.upload(file=pdf_path)
            while uploaded.state == types.FileState.PROCESSING:
                time.sleep(1)
                uploaded = client.files.get(name=uploaded.name)
            if uploaded.state != types.FileState.ACTIVE:
                raise RuntimeError(f"وضعیت غیرمنتظره فایل در سرور گوگل: {uploaded.state}")
            uploaded_name = uploaded.name
            contents = [uploaded]

        response = client.models.generate_content(
            model=model_name,
            contents=contents,
            config=types.GenerateContentConfig(
                system_instruction=EXTRACTION_PROMPT,
                temperature=0,
            ),
        )
        raw_text = response.text or ""
    finally:
        if uploaded_name:
            try:
                client.files.delete(name=uploaded_name)
            except Exception:
                pass

    cleaned_lines = [ln for ln in raw_text.split('\n') if not ln.strip().startswith("```")]
    return "\n".join(cleaned_lines).strip()


# ==============================================================================
# منطق اصلی پردازش (در ترد پس‌زمینه اجرا می‌شود)
# ==============================================================================

def process_workflow(params, log_fn, progress_fn, stop_event):
    pdf_path = params["pdf_path"]
    base_name = params["base_name"]
    pages = params["pages"]
    output_dir = params["output_dir"]
    model_name = params["model_name"]
    api_key = params["api_key"]
    resume = params["resume"]

    project_folder = os.path.join(output_dir, base_name)
    os.makedirs(project_folder, exist_ok=True)
    word_path = os.path.join(project_folder, f"{base_name}.docx")

    completed = set()
    if resume and os.path.exists(word_path):
        completed = scan_completed_pages(word_path)
        doc = Document(word_path)
        log_fn(f"[INFO] ادامه از فایل قبلی؛ {len(completed)} صفحه قبلاً کامل شده بود و رد می‌شود.")
    else:
        doc = Document()
        title_p = doc.add_paragraph()
        set_rtl_paragraph(title_p)
        title_run = title_p.add_run(f"سوالات استخراج‌شده — پروژه: {base_name}")
        apply_persian_font(title_run, size=15, bold=True)
        sub_p = doc.add_paragraph()
        set_rtl_paragraph(sub_p)
        sub_run = sub_p.add_run(f"منبع: {os.path.basename(pdf_path)}")
        apply_persian_font(sub_run, size=10)

    target_pages = [p for p in pages if p not in completed]
    total = len(pages)
    done_count = len(completed)
    progress_fn(done_count, total)

    if not target_pages:
        log_fn("[INFO] همه صفحات درخواستی قبلاً استخراج شده بودند؛ کاری برای انجام نیست.")
        return {"word_path": word_path, "ok": True}

    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        log_fn(f"[FATAL] باز کردن فایل PDF ناموفق بود: {e}")
        return {"word_path": word_path, "ok": False}

    total_pdf_pages = len(reader.pages)
    client = genai.Client(api_key=api_key)

    any_failure = False
    fatal_stop = False

    for i, page_num in enumerate(target_pages):
        if stop_event.is_set():
            log_fn("[INFO] پردازش توسط کاربر متوقف شد.")
            break

        idx = page_num - 1
        if idx < 0 or idx >= total_pdf_pages:
            log_fn(f"[WARN] صفحه {page_num} در فایل PDF وجود ندارد (تعداد کل صفحات: {total_pdf_pages}). رد شد.")
            done_count += 1
            progress_fn(done_count, total)
            continue

        writer = PdfWriter()
        writer.add_page(reader.pages[idx])
        single_filename = f"{base_name}(page{page_num}).pdf"
        single_path = os.path.join(project_folder, single_filename)
        with open(single_path, "wb") as f:
            writer.write(f)

        log_fn(f"[INFO] فایل {single_filename} ساخته شد. در حال ارسال به {model_name}...")

        extracted_text = None
        last_error = None

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                extracted_text = call_gemini_for_page(client, model_name, single_path)
                break
            except genai_errors.ClientError as e:
                last_error = f"{e.code} {e.message}"
                if e.code in FATAL_STOP_CODES:
                    if e.code == 404:
                        log_fn(f"[FATAL] مدل «{model_name}» پیدا نشد (404). نام مدل را در برنامه بررسی/تغییر بده.")
                    else:
                        log_fn(f"[FATAL] خطای احراز هویت (کد {e.code}): {e.message}\nکلید API را بررسی یا جایگزین کن.")
                    fatal_stop = True
                    break
                elif e.code == 429:
                    wait = BASE_BACKOFF_SECONDS * (2 ** (attempt - 1))
                    log_fn(f"[RETRY] محدودیت نرخ درخواست (429) — تلاش {attempt}/{MAX_RETRIES}. {wait} ثانیه صبر...")
                    if attempt < MAX_RETRIES:
                        time.sleep(wait)
                else:
                    log_fn(f"[ERROR] تلاش {attempt}/{MAX_RETRIES} برای صفحه {page_num} ناموفق: {e.message}")
                    if attempt < MAX_RETRIES:
                        time.sleep(BASE_BACKOFF_SECONDS)
            except genai_errors.ServerError as e:
                last_error = f"{e.code} {e.message}"
                wait = BASE_BACKOFF_SECONDS * (2 ** (attempt - 1))
                log_fn(f"[RETRY] خطای موقت سرور گوگل (کد {e.code}) — تلاش {attempt}/{MAX_RETRIES}. {wait} ثانیه صبر...")
                if attempt < MAX_RETRIES:
                    time.sleep(wait)
            except Exception as e:
                last_error = str(e)
                log_fn(f"[ERROR] تلاش {attempt}/{MAX_RETRIES} برای صفحه {page_num} ناموفق: {e}")
                if attempt < MAX_RETRIES:
                    time.sleep(BASE_BACKOFF_SECONDS)

        remove_stale_error_paragraph(doc, page_num)

        header_p = doc.add_paragraph()
        set_rtl_paragraph(header_p)

        if extracted_text is not None:
            header_run = header_p.add_run(f"--- سوالات صفحه {page_num} ---")
            apply_persian_font(header_run, size=13, bold=True)
            for line in extracted_text.split('\n'):
                p_line = doc.add_paragraph()
                set_rtl_paragraph(p_line)
                r = p_line.add_run(line)
                apply_persian_font(r, size=12)
            log_fn(f"[SUCCESS] صفحه {page_num} با موفقیت استخراج شد.")
        else:
            any_failure = True
            header_run = header_p.add_run(f"--- خطا در استخراج سوالات صفحه {page_num} ---")
            apply_persian_font(header_run, size=13, bold=True)
            log_fn(f"[FATAL] صفحه {page_num} پس از {MAX_RETRIES} تلاش ناموفق بود. آخرین خطا: {last_error}")

        try:
            doc.save(word_path)
        except Exception as e:
            log_fn(f"[WARN] ذخیره خودکار فایل Word ناموفق بود (ادامه می‌دهیم): {e}")

        done_count += 1
        progress_fn(done_count, total)

        if fatal_stop:
            log_fn("[INFO] چون این خطا برای همه صفحات دیگر هم تکرار می‌شود، پردازش کامل متوقف شد.")
            break

        is_last = (i == len(target_pages) - 1)
        if not is_last and not stop_event.is_set():
            time.sleep(REQUEST_DELAY_SECONDS)

    try:
        doc.save(word_path)
    except Exception as e:
        log_fn(f"[ERROR] ذخیره نهایی فایل Word ناموفق بود: {e}")
        return {"word_path": word_path, "ok": False}

    return {"word_path": word_path, "ok": not (any_failure or fatal_stop)}


# ==============================================================================
# رابط کاربری گرافیکی (Tkinter)
# ==============================================================================

class ExtractionApp:
    def __init__(self, root):
        self.root = root
        root.title("استخراج سوالات از PDF — Gemini")
        root.geometry("760x620")
        root.minsize(680, 560)

        self.default_font = ("Tahoma", 10)
        root.option_add("*Font", self.default_font)

        self.pdf_path_var = tk.StringVar()
        self.base_name_var = tk.StringVar()
        self.pages_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.model_var = tk.StringVar(value=MODEL_NAME_DEFAULT)
        self.status_var = tk.StringVar(value="آماده")

        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.last_project_folder = None

        self._build_ui()
        self.root.after(100, self._poll_queue)

    # ---------------------------------------------------------------- UI ----
    def _build_ui(self):
        pad = {"padx": 8, "pady": 5}

        form = ttk.LabelFrame(self.root, text="تنظیمات پروژه")
        form.pack(fill="x", padx=10, pady=10)
        form.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(form, text="فایل PDF ورودی:").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(form, textvariable=self.pdf_path_var, justify="right", state="readonly").grid(
            row=row, column=1, sticky="ew", **pad)
        ttk.Button(form, text="انتخاب فایل...", command=self.on_browse_pdf).grid(row=row, column=2, **pad)

        row += 1
        ttk.Label(form, text="نام پروژه (مثلاً Bio10.3):").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(form, textvariable=self.base_name_var, justify="right").grid(
            row=row, column=1, columnspan=2, sticky="ew", **pad)

        row += 1
        ttk.Label(form, text="محدوده صفحات:").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(form, textvariable=self.pages_var, justify="right").grid(
            row=row, column=1, columnspan=2, sticky="ew", **pad)

        row += 1
        ttk.Label(form, text="نمونه: 250-255, 260, 270-275", foreground="#666").grid(
            row=row, column=1, columnspan=2, sticky="w", padx=8)

        row += 1
        ttk.Label(form, text="پوشه ذخیره‌سازی:").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(form, textvariable=self.output_dir_var, justify="right", state="readonly").grid(
            row=row, column=1, sticky="ew", **pad)
        ttk.Button(form, text="انتخاب پوشه...", command=self.on_browse_dir).grid(row=row, column=2, **pad)

        row += 1
        ttk.Label(form, text="مدل Gemini:").grid(row=row, column=0, sticky="e", **pad)
        model_combo = ttk.Combobox(
            form, textvariable=self.model_var, justify="right",
            values=["gemini-3.1-flash-lite", "gemini-3.5-flash-lite"],
        )
        model_combo.grid(row=row, column=1, columnspan=2, sticky="ew", **pad)

        # buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.start_btn = ttk.Button(btn_frame, text="شروع", command=self.on_start)
        self.start_btn.pack(side="right", padx=4)
        self.stop_btn = ttk.Button(btn_frame, text="توقف", command=self.on_stop, state="disabled")
        self.stop_btn.pack(side="right", padx=4)
        self.open_folder_btn = ttk.Button(btn_frame, text="باز کردن پوشه خروجی", command=self.on_open_folder,
                                           state="disabled")
        self.open_folder_btn.pack(side="right", padx=4)

        # progress
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True, padx=(0, 8))
        ttk.Label(prog_frame, textvariable=self.status_var, width=16).pack(side="right")

        # log
        log_frame = ttk.LabelFrame(self.root, text="گزارش زنده")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, wrap="word", height=14, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    # ----------------------------------------------------------- actions ----
    def on_browse_pdf(self):
        path = filedialog.askopenfilename(title="فایل PDF اصلی را انتخاب کنید", filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path_var.set(path)

    def on_browse_dir(self):
        d = filedialog.askdirectory(title="پوشه محل ذخیره پروژه را انتخاب کنید")
        if d:
            self.output_dir_var.set(d)

    def on_open_folder(self):
        if self.last_project_folder and os.path.isdir(self.last_project_folder):
            try:
                if sys.platform == "win32":
                    os.startfile(self.last_project_folder)
                else:
                    messagebox.showinfo("مسیر پوشه", self.last_project_folder)
            except Exception as e:
                messagebox.showwarning("خطا", f"باز کردن پوشه ناموفق بود: {e}")

    def _append_log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _load_or_prompt_api_key(self):
        load_dotenv(ENV_PATH)
        key = os.environ.get("GEMINI_API_KEY", "").strip()
        if key:
            return key

        key = simpledialog.askstring(
            "کلید Gemini API",
            "کلید Gemini API را وارد کنید.\n(از aistudio.google.com/apikey قابل دریافت است):",
            show="*", parent=self.root,
        )
        if not key:
            return None
        key = key.strip()

        if messagebox.askyesno(
            "ذخیره کلید",
            "این کلید در فایل .env کنار برنامه ذخیره شود تا دفعه بعد لازم نباشد دوباره وارد کنی؟\n"
            "(این فایل را با کسی به اشتراک نگذار)"
        ):
            try:
                if not ENV_PATH.exists():
                    ENV_PATH.touch()
                set_key(str(ENV_PATH), "GEMINI_API_KEY", key)
            except Exception as e:
                messagebox.showwarning("هشدار", f"ذخیره کلید در .env ناموفق بود: {e}")

        os.environ["GEMINI_API_KEY"] = key
        return key

    def on_start(self):
        pdf_path = self.pdf_path_var.get().strip()
        base_name = self.base_name_var.get().strip()
        pages_str = self.pages_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        model_name = self.model_var.get().strip() or MODEL_NAME_DEFAULT

        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showerror("خطا", "لطفاً یک فایل PDF معتبر انتخاب کن.")
            return
        if not base_name or not is_valid_name_component(base_name):
            messagebox.showerror("خطا", 'نام پروژه نامعتبر است (نباید شامل کاراکترهای  < > : " / \\ | ? *  باشد).')
            return
        if not output_dir or not os.path.isdir(output_dir):
            messagebox.showerror("خطا", "لطفاً پوشه ذخیره‌سازی را انتخاب کن.")
            return
        try:
            pages = parse_pages_input(pages_str)
        except Exception:
            messagebox.showerror("خطا", "فرمت محدوده صفحات نادرست است.\nنمونه صحیح: 250-255, 260, 270-275")
            return

        project_folder = os.path.join(output_dir, base_name)
        word_path = os.path.join(project_folder, f"{base_name}.docx")
        resume = False
        if os.path.exists(word_path):
            resume = messagebox.askyesno(
                "فایل موجود",
                f"یک فایل Word با نام «{base_name}.docx» از قبل در این پوشه هست.\n\n"
                "ادامه بده (فقط صفحات باقی‌مانده اضافه شود)؟\n"
                "در صورت انتخاب «خیر»، این فایل از نو ساخته می‌شود."
            )

        api_key = self._load_or_prompt_api_key()
        if not api_key:
            return

        self.last_project_folder = project_folder
        self.log_text.delete("1.0", "end")
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.open_folder_btn.config(state="normal")
        self.status_var.set("در حال پردازش...")
        self.stop_event.clear()

        params = dict(pdf_path=pdf_path, base_name=base_name, pages=pages, output_dir=output_dir,
                      model_name=model_name, api_key=api_key, resume=resume)

        self.worker_thread = threading.Thread(target=self._run_worker, args=(params,), daemon=True)
        self.worker_thread.start()

    def on_stop(self):
        self.stop_event.set()
        self._append_log("[INFO] درخواست توقف ثبت شد؛ پس از پایان صفحه فعلی، پردازش متوقف می‌شود...")
        self.stop_btn.config(state="disabled")

    def _run_worker(self, params):
        result = {"ok": False, "word_path": None}
        try:
            result = process_workflow(
                params,
                log_fn=lambda m: self.log_queue.put(("log", m)),
                progress_fn=lambda d, t: self.log_queue.put(("progress", (d, t))),
                stop_event=self.stop_event,
            )
        except Exception:
            self.log_queue.put(("log", f"[FATAL] خطای غیرمنتظره:\n{traceback.format_exc()}"))
        finally:
            self.log_queue.put(("done", result))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "progress":
                    done, total = payload
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = done
                    self.status_var.set(f"{done} از {total} صفحه")
                elif kind == "done":
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    if payload.get("ok"):
                        self.status_var.set("پایان یافت")
                        messagebox.showinfo("موفقیت", f"پردازش تمام شد.\nفایل Word:\n{payload.get('word_path')}")
                    else:
                        self.status_var.set("پایان با خطا")
                        messagebox.showwarning(
                            "پایان با خطا",
                            f"پردازش تمام شد اما برخی صفحات با خطا مواجه شدند یا متوقف شد.\n"
                            f"جزئیات را در گزارش زنده ببین.\nفایل Word:\n{payload.get('word_path')}"
                        )
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)


def main():
    root = tk.Tk()
    ExtractionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
