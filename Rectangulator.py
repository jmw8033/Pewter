from email.mime.multipart import MIMEMultipart
from matplotlib.widgets import TextBox
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
from matplotlib.widgets import CheckButtons
from email.mime.text import MIMEText
from Alertinator import AlertWindow
from datetime import datetime
from difflib import SequenceMatcher
from collections import OrderedDict
import tkinter as tk
from tkinter import filedialog
import numpy as np
import traceback
import keyring
import warnings
import smtplib
import json
import fitz
import glob
import sys
import re
import os

warnings.simplefilter("ignore", UserWarning)

if getattr(sys, 'frozen', False):
    # If running as a compiled .exe, use the folder containing the .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # If running as a normal Python script, use the folder containing the script
    BASE_DIR = os.path.dirname(__file__)

with open(os.path.join(BASE_DIR, "config.json")) as f:
    config = json.load(f)

# Backward-compatible defaults. EmailProcessor persists missing values to config.json.
_template_parent = os.path.dirname(str(config.get("TEMPLATE_FOLDER", "") or "")) or BASE_DIR
_test_template_parent = os.path.dirname(str(config.get("TEST_TEMPLATE_FOLDER", "") or "")) or BASE_DIR
config.setdefault("OCR_TEMPLATE_FOLDER", os.path.join(_template_parent, "OCR_Templates"))
config.setdefault("STATEMENT_TEMPLATE_FOLDER", os.path.join(_template_parent, "Statement_Templates"))
_invoice_parent = os.path.dirname(str(config.get("INVOICE_FOLDER", "") or "")) or BASE_DIR
_test_invoice_parent = os.path.dirname(str(config.get("TEST_INVOICE_FOLDER", "") or "")) or BASE_DIR
config.setdefault("STATEMENT_FOLDER", os.path.join(_invoice_parent, "Statements"))
config.setdefault("TEST_OCR_TEMPLATE_FOLDER", os.path.join(_test_template_parent, "OCR_Templates"))
config.setdefault("TEST_STATEMENT_TEMPLATE_FOLDER", os.path.join(_test_template_parent, "Statement_Templates"))
config.setdefault("TEST_STATEMENT_FOLDER", os.path.join(_test_invoice_parent, "Statements"))
config.setdefault("OCR_FUZZY_THRESHOLD", 0.72)
config.setdefault("OCR_DPI", 250)
config.setdefault("OCR_LANGUAGE", "eng")
config.setdefault("TESSDATA_PREFIX", "")
config.setdefault("PRINTER_NAME", "")
config.setdefault("MIN_EMBEDDED_TEXT_CHARS", 40)
config.setdefault("POSTFIX_VENDORS", "")

def get_password(username):
    password = keyring.get_password("PewterInvoiceProcessor", username)
    if password is None:
        print(f"No stored password for '{username}'. Store one with:")
        print(f"  keyring.set_password('PewterInvoiceProcessor', '{username}', '<app password>')")
        exit(1)
    return password

class RectangulatorHandler:

    # Rendering quality knobs. The visible region is re-rendered by PyMuPDF at
    # roughly the canvas's own pixel resolution, so text stays sharp no matter
    # how far you zoom in. SUPERSAMPLE > 1 renders extra detail so small zoom
    # steps don't each need a fresh render.
    SUPERSAMPLE = 2.0
    MAX_RENDER_ZOOM = 16.0
    MAX_RENDER_PIXELS = 16_000_000

    def __init__(self, root, fig, ax):
        self.queue = []
        self.invoice = True
        self.pdf_path = None
        self.page_rect = None
        self.page_image = None
        self.current_page = 0
        self.total_pages = 1
        self._render_job = None
        self.should_print = True
        self.should_save = True
        self.hit_submit = False
        self.root = root
        self.fig = fig
        self.ax = ax
        self.done_var = tk.IntVar()
        self.config = dict(config)
        self.palette = getattr(root, "palette", {
            "surface": "#FFFFFF", "surface_alt": "#F5F7F8", "border": "#D4DDE1",
            "text": "#26343A", "muted": "#6B7A81", "pewter": "#77878E",
            "pewter_dark": "#4F6067", "accent": "#3E7180", "accent_hover": "#335E6A",
            "success": "#3E7F66", "success_soft": "#E4F2EB",
            "warning": "#B67C28", "warning_soft": "#FFF1D7",
            "danger": "#B55353", "danger_soft": "#FBE7E7", "purple": "#766A91",
        })
        # Field colors deliberately mirror the order the user selects them:
        # vendor → date → invoice number. This makes correction/retraining easier
        # to understand without adding more on-screen instructions.
        self.field_colors = [self.palette["accent"], self.palette["warning"], self.palette["purple"]]

        # Small LRU caches: native text and OCR are both expensive enough that
        # reopening/re-reading a PDF for every template is wasteful.
        self._document_cache = OrderedDict()
        self._document_cache_limit = 8
        self._template_cache = {}
        self._ocr_template_cache = {}
        self._statement_template_cache = {}
        self._ocr_errors_reported = set()

        # State used by the human-confirmed OCR workflow.
        self.ocr_mode = False
        self.ocr_template_ready = False
        self.ocr_match = None
        self.statement_mode = False
        self.statement_match = None
        self.statement_company = ""
        self.statement_destination_folder = ""
        self.statement_root = ""
        self.statement_template_folder = ""
        self.statement_ocr_mode = False

    def refresh_config(self):
        with open(os.path.join(BASE_DIR, "config.json"), "r", encoding="utf-8") as f:
            self.config = json.load(f)
        template_parent = os.path.dirname(str(self.config.get("TEMPLATE_FOLDER", "") or "")) or BASE_DIR
        test_template_parent = os.path.dirname(str(self.config.get("TEST_TEMPLATE_FOLDER", "") or "")) or BASE_DIR
        self.config.setdefault("OCR_TEMPLATE_FOLDER", os.path.join(template_parent, "OCR_Templates"))
        self.config.setdefault("STATEMENT_TEMPLATE_FOLDER", os.path.join(template_parent, "Statement_Templates"))
        invoice_parent = os.path.dirname(str(self.config.get("INVOICE_FOLDER", "") or "")) or BASE_DIR
        test_invoice_parent = os.path.dirname(str(self.config.get("TEST_INVOICE_FOLDER", "") or "")) or BASE_DIR
        self.config.setdefault("STATEMENT_FOLDER", os.path.join(invoice_parent, "Statements"))
        self.config.setdefault("TEST_OCR_TEMPLATE_FOLDER", os.path.join(test_template_parent, "OCR_Templates"))
        self.config.setdefault("TEST_STATEMENT_TEMPLATE_FOLDER", os.path.join(test_template_parent, "Statement_Templates"))
        self.config.setdefault("TEST_STATEMENT_FOLDER", os.path.join(test_invoice_parent, "Statements"))
        self.config.setdefault("OCR_FUZZY_THRESHOLD", 0.72)
        self.config.setdefault("OCR_DPI", 250)
        self.config.setdefault("OCR_LANGUAGE", "eng")
        self.config.setdefault("TESSDATA_PREFIX", "")
        self.config.setdefault("PRINTER_NAME", "")
        self.config.setdefault("MIN_EMBEDDED_TEXT_CHARS", 40)
        self.config.setdefault("POSTFIX_VENDORS", "")
        self._template_cache.clear()
        self._ocr_template_cache.clear()
        self._statement_template_cache.clear()

    def _document_signature(self, pdf_path):
        try:
            stat = os.stat(pdf_path)
            return (stat.st_mtime_ns, stat.st_size)
        except OSError:
            return None

    def _cache_entry(self, pdf_path):
        path = os.path.abspath(pdf_path)
        signature = self._document_signature(path)
        entry = self._document_cache.get(path)
        if entry is None or entry.get("signature") != signature:
            entry = {"signature": signature, "native": {}, "ocr": {}, "page_rects": {}}
            self._document_cache[path] = entry
        else:
            self._document_cache.move_to_end(path)
        while len(self._document_cache) > self._document_cache_limit:
            self._document_cache.popitem(last=False)
        return entry

    def invalidate_document_cache(self, pdf_path=None):
        if pdf_path is None:
            self._document_cache.clear()
            return
        self._document_cache.pop(os.path.abspath(pdf_path), None)

    def invalidate_template_cache(self, folder=None):
        if folder is None:
            self._template_cache.clear()
            self._ocr_template_cache.clear()
            self._statement_template_cache.clear()
        else:
            folder = os.path.abspath(folder)
            self._template_cache.pop(folder, None)
            self._ocr_template_cache.pop(folder, None)
            self._statement_template_cache.pop(folder, None)

    @staticmethod
    def _normalise_selection(rect):
        if hasattr(rect, "get_x"):
            x0 = float(rect.get_x())
            y0 = float(rect.get_y())
            x1 = x0 + float(rect.get_width())
            y1 = y0 + float(rect.get_height())
        elif isinstance(rect, fitz.Rect):
            x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
        else:
            x0, y0, x1, y1 = map(float, rect)
        return fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

    def _page_rect(self, pdf_path, page_num):
        entry = self._cache_entry(pdf_path)
        if page_num in entry["page_rects"]:
            return fitz.Rect(entry["page_rects"][page_num])
        with fitz.open(pdf_path) as doc:
            if not doc:
                raise ValueError("PDF contains no pages")
            page_num = min(max(int(page_num), 0), len(doc) - 1)
            rect = fitz.Rect(doc[page_num].rect)
        entry["page_rects"][page_num] = tuple(rect)
        return rect

    def _get_page_words(self, pdf_path, page_num=0, ocr=False):
        entry = self._cache_entry(pdf_path)
        cache_name = "ocr" if ocr else "native"
        cache = entry[cache_name]
        page_num = int(page_num)
        if page_num in cache:
            return cache[page_num]

        with fitz.open(pdf_path) as doc:
            if not doc:
                return []
            page_num = min(max(page_num, 0), len(doc) - 1)
            page = doc[page_num]
            entry["page_rects"][page_num] = tuple(page.rect)
            if not ocr:
                words = page.get_text("words", sort=True)
            else:
                tessdata = str(self.config.get("TESSDATA_PREFIX", "") or "").strip()
                if tessdata:
                    os.environ["TESSDATA_PREFIX"] = tessdata
                try:
                    textpage = page.get_textpage_ocr(
                        language=str(self.config.get("OCR_LANGUAGE", "eng") or "eng"),
                        dpi=int(self.config.get("OCR_DPI", 250) or 250),
                        full=True,
                    )
                    words = page.get_text("words", textpage=textpage, sort=True)
                except Exception as exc:
                    key = (type(exc).__name__, str(exc))
                    if key not in self._ocr_errors_reported:
                        self._ocr_errors_reported.add(key)
                        self.log(
                            "OCR is unavailable. Install Tesseract and, if needed, set "
                            f"TESSDATA_PREFIX in Settings. Details: {exc}",
                            tag="red", display=True,
                        )
                    words = []
        cache[page_num] = words
        return words

    def document_has_embedded_text(self, pdf_path, minimum_chars=None):
        try:
            if minimum_chars is None:
                minimum_chars = int(self.config.get("MIN_EMBEDDED_TEXT_CHARS", 40) or 40)
            with fitz.open(pdf_path) as doc:
                pages = len(doc)
            count = 0
            for page_num in range(pages):
                for word in self._get_page_words(pdf_path, page_num, ocr=False):
                    count += len(str(word[4]).strip())
                    if count >= minimum_chars:
                        return True
            return False
        except Exception as exc:
            self.log(f"Unable to inspect PDF text layer: {exc}", tag="orange", display=True)
            return False

    def get_page_text(self, pdf_path, page_num=0, ocr=False):
        return " ".join(str(word[4]) for word in self._get_page_words(pdf_path, page_num, ocr=ocr)).strip()

    def get_full_text(self, pdf_path, ocr=False):
        try:
            with fitz.open(pdf_path) as doc:
                pages = len(doc)
            return "\n".join(self.get_page_text(pdf_path, i, ocr=ocr) for i in range(pages))
        except Exception:
            return ""

    def _field_rect(self, field, pdf_path):
        page_num = int(field.get("page", 0))
        rect = field.get("rect", [0, 0, 0, 0])
        if field.get("normalized", False):
            page_rect = self._page_rect(pdf_path, page_num)
            x, y, w, h = map(float, rect)
            return Rectangle(
                (page_rect.x0 + x * page_rect.width, page_rect.y0 + y * page_rect.height),
                w * page_rect.width,
                h * page_rect.height,
            ), page_num
        x, y, w, h = map(float, rect)
        return Rectangle((x, y), w, h), page_num

    def normalized_field(self, rect, pdf_path, page_num, expected_text=""):
        selection = self._normalise_selection(rect)
        page_rect = self._page_rect(pdf_path, page_num)
        if page_rect.width <= 0 or page_rect.height <= 0:
            coords = [selection.x0, selection.y0, selection.width, selection.height]
            normalized = False
        else:
            coords = [
                (selection.x0 - page_rect.x0) / page_rect.width,
                (selection.y0 - page_rect.y0) / page_rect.height,
                selection.width / page_rect.width,
                selection.height / page_rect.height,
            ]
            normalized = True
        return {
            "page": int(page_num),
            "rect": [round(float(v), 8) for v in coords],
            "normalized": normalized,
            "expected": str(expected_text),
        }

    @staticmethod
    def _folder_signature(folder, patterns=("*.txt", "*.json")):
        files = []
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(folder, pattern)))
        signature = []
        for path in sorted(set(files)):
            try:
                st = os.stat(path)
                signature.append((path, st.st_mtime_ns, st.st_size))
            except OSError:
                pass
        return tuple(signature)

    def _load_native_templates(self, folder):
        folder = os.path.abspath(folder)
        os.makedirs(folder, exist_ok=True)
        signature = self._folder_signature(folder)
        cached = self._template_cache.get(folder)
        if cached and cached[0] == signature:
            return cached[1]

        templates = []
        # Existing delimiter-based templates remain fully supported; they are
        # treated as page 1 because the old format did not store page numbers.
        for path in glob.glob(os.path.join(folder, "*.txt")):
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    lines = [line.rstrip("\n") for line in f]
                for i in range(0, len(lines) - 2, 3):
                    parsed = []
                    for line in lines[i:i+3]:
                        parts = line.split("?")
                        if len(parts) < 5:
                            parsed = []
                            break
                        parsed.append({
                            "expected": parts[0],
                            "page": 0,
                            "rect": [float(parts[1]), float(parts[2]), float(parts[3]), float(parts[4])],
                            "normalized": False,
                        })
                    if parsed:
                        templates.append({
                            "source": path,
                            "vendor": parsed[0]["expected"],
                            "fields": {"vendor": parsed[0], "date": parsed[1], "invoice_number": parsed[2]},
                        })
            except Exception as exc:
                self.log(f"Error loading template {path}: {exc}", tag="red", display=True)

        for path in glob.glob(os.path.join(folder, "*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                if payload.get("type") != "text":
                    continue
                variants = payload.get("templates") or [payload]
                for variant in variants:
                    fields = variant.get("fields", {})
                    if all(k in fields for k in ("vendor", "date", "invoice_number")):
                        templates.append({
                            "source": path,
                            "vendor": variant.get("vendor", payload.get("vendor", fields["vendor"].get("expected", ""))),
                            "fields": fields,
                        })
            except Exception as exc:
                self.log(f"Error loading template {path}: {exc}", tag="red", display=True)
        self._template_cache[folder] = (signature, templates)
        return templates

    def _load_ocr_templates(self, folder):
        folder = os.path.abspath(folder)
        os.makedirs(folder, exist_ok=True)
        signature = self._folder_signature(folder, patterns=("*.json",))
        cached = self._ocr_template_cache.get(folder)
        if cached and cached[0] == signature:
            return cached[1]
        templates = []
        for path in glob.glob(os.path.join(folder, "*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                if payload.get("type") != "ocr":
                    continue
                variants = payload.get("templates") or [payload]
                for variant in variants:
                    fields = variant.get("fields", {})
                    if all(k in fields for k in ("vendor", "date", "invoice_number")):
                        templates.append({
                            "source": path,
                            "vendor": variant.get("vendor", payload.get("vendor", fields["vendor"].get("expected", ""))),
                            "fields": fields,
                        })
            except Exception as exc:
                self.log(f"Error loading OCR template {path}: {exc}", tag="red", display=True)
        self._ocr_template_cache[folder] = (signature, templates)
        return templates

    @staticmethod
    def _normalise_match_text(value):
        value = re.sub(r"[^A-Za-z0-9]+", " ", str(value or "")).casefold()
        return " ".join(value.split())

    def fuzzy_score(self, expected, actual):
        expected = self._normalise_match_text(expected)
        actual = self._normalise_match_text(actual)
        if not expected or not actual:
            return 0.0
        direct = SequenceMatcher(None, expected, actual).ratio()
        token_expected = " ".join(sorted(expected.split()))
        token_actual = " ".join(sorted(actual.split()))
        token = SequenceMatcher(None, token_expected, token_actual).ratio()
        return max(direct, token)

    def build_filename_from_fields(self, vendor, invoice_date, invoice_num):
        vendor_clean = self.sanitize_filename(vendor)
        prefix = self.get_vendor_prefix(vendor_clean)
        postfix = self.get_vendor_postfix(vendor_clean)
        invoice_num = str(invoice_num or "").strip()
        if prefix and not invoice_num.startswith(prefix):
            invoice_num = f"{prefix}{invoice_num}"
        if postfix and not invoice_num.endswith(postfix):
            invoice_num = f"{invoice_num}{postfix}"
        invoice_date = self.clean_date(str(invoice_date or "").strip())
        return self.sanitize_filename(f"{invoice_date}_{invoice_num}")

    def suggest_invoice_filename(self, pdf_path, ocr=False):
        # Lightweight fallback suggestion for unknown text PDFs. It never saves
        # automatically; it merely pre-fills the manual text box.
        text = self.get_full_text(pdf_path, ocr=ocr)
        if not text:
            return ""
        date_match = re.search(
            r"(?:invoice\s+date|date)\s*[:#-]?\s*"
            r"([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}|\d{1,4}[/-]\d{1,2}[/-]\d{1,4})",
            text, re.I,
        )
        num_match = re.search(
            r"(?:invoice\s*(?:number|no\.?|#)|inv\.?\s*(?:no\.?|#))\s*[:#-]?\s*([A-Za-z0-9._/-]+)",
            text, re.I,
        )
        if not date_match or not num_match:
            return ""
        return self.sanitize_filename(f"{self.clean_date(date_match.group(1))}_{num_match.group(1)}")

    def _load_statement_templates(self, folder):
        folder = os.path.abspath(folder)
        os.makedirs(folder, exist_ok=True)
        signature = self._folder_signature(folder, patterns=("*.json",))
        cached = self._statement_template_cache.get(folder)
        if cached and cached[0] == signature:
            return cached[1]
        templates = []
        for path in glob.glob(os.path.join(folder, "*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                if payload.get("type") != "statement":
                    continue
                variants = payload.get("templates") or [payload]
                for variant in variants:
                    fields = variant.get("fields", {})
                    if all(k in fields for k in ("identifier", "date")):
                        templates.append({
                            "source": path,
                            "company": variant.get("company", payload.get("company", "")),
                            "folder": variant.get("folder", payload.get("folder", "")),
                            "ocr": bool(variant.get("ocr", payload.get("ocr", False))),
                            "fields": fields,
                        })
            except Exception as exc:
                self.log(f"Error loading statement template {path}: {exc}", tag="red", display=True)
        self._statement_template_cache[folder] = (signature, templates)
        return templates

    def check_statement_templates(self, pdf_path, template_folder, statement_root, ocr=False):
        """Return the best statement template for the document, if one matches."""
        templates = self._load_statement_templates(template_folder)
        best = None
        threshold = float(self.config.get("OCR_FUZZY_THRESHOLD", 0.72) or 0.72)
        for template in templates:
            if bool(template.get("ocr")) != bool(ocr):
                continue
            field = template["fields"]["identifier"]
            rect, page_num = self._field_rect(field, pdf_path)
            actual = self.get_text_in_rect(rect, pdf_path, page_num, ocr=ocr)
            expected = field.get("expected", "")
            if ocr:
                score = self.fuzzy_score(expected, actual)
            else:
                score = 1.0 if self._normalise_match_text(expected) == self._normalise_match_text(actual) and actual.strip() else 0.0
            if best is None or score > best["score"]:
                best = {"score": score, "template": template, "actual_identifier": actual}
        required = threshold if ocr else 1.0
        if best is None or best["score"] < required:
            return None

        template = best["template"]
        date_field = template["fields"]["date"]
        date_rect, date_page = self._field_rect(date_field, pdf_path)
        statement_date = self.get_text_in_rect(date_rect, pdf_path, date_page, ocr=ocr)
        company = self.sanitize_filename(template.get("company", ""))
        rel_folder = str(template.get("folder", "") or "").strip()
        root_abs = os.path.abspath(statement_root)
        destination = os.path.abspath(os.path.join(root_abs, rel_folder)) if rel_folder else os.path.join(root_abs, company)
        try:
            if os.path.commonpath([root_abs, destination]) != root_abs:
                return None
        except ValueError:
            return None
        best.update({
            "company": company,
            "destination_folder": destination,
            "statement_date": statement_date,
            "suggested_filename": self.build_statement_filename(company, statement_date),
        })
        return best

    def build_statement_filename(self, company, statement_date):
        company = self.sanitize_filename(company)
        date_text = self.clean_date(str(statement_date or "").strip())
        return self.sanitize_filename(f"{company} Statement {date_text}")

    def statement_destination(self, entered_filename):
        entered = self.sanitize_filename(entered_filename)
        if not entered or not self.statement_destination_folder:
            return None
        # Convenience: if the user replaces the whole bar with only a date,
        # restore the standard Company Statement Date naming convention.
        date_candidate = self.clean_date(entered)
        if self.statement_company and date_candidate != entered:
            entered = self.build_statement_filename(self.statement_company, entered)
        return os.path.join(self.statement_destination_folder, f"{entered}.pdf")

    def check_ocr_templates(self, pdf_path, template_folder):
        templates = self._load_ocr_templates(template_folder)
        if not templates:
            return None

        best = None
        for template in templates:  # deliberately score every template
            field = template["fields"]["vendor"]
            rect, page_num = self._field_rect(field, pdf_path)
            actual = self.get_text_in_rect(rect, pdf_path, page_num, ocr=True)
            expected = field.get("expected") or template.get("vendor", "")
            score = self.fuzzy_score(expected, actual)
            if best is None or score > best["score"]:
                best = {"score": score, "template": template, "actual_vendor": actual}

        threshold = float(self.config.get("OCR_FUZZY_THRESHOLD", 0.72) or 0.72)
        if best is None or best["score"] < threshold:
            return None

        fields = best["template"]["fields"]
        date_rect, date_page = self._field_rect(fields["date"], pdf_path)
        num_rect, num_page = self._field_rect(fields["invoice_number"], pdf_path)
        invoice_date = self.get_text_in_rect(date_rect, pdf_path, date_page, ocr=True)
        invoice_num = self.get_text_in_rect(num_rect, pdf_path, num_page, ocr=True)
        vendor = fields["vendor"].get("expected") or best["template"].get("vendor", best["actual_vendor"])
        best["invoice_date"] = invoice_date
        best["invoice_number"] = invoice_num
        best["suggested_filename"] = self.build_filename_from_fields(vendor, invoice_date, invoice_num)
        return best

    def rectangulate(self, filename, filepath, root, template_folder, testing=False,
                    ocr_mode=False, ocr_template_folder=None, statement_folder=None,
                    statement_template_folder=None):
        # Away mode intentionally bypasses review, matching the existing behavior.
        if root.AWAY_MODE and not testing:
            self.print_invoice(filepath)
            return [filepath, False]

        self.ocr_mode = bool(ocr_mode)
        self.ocr_match = None
        self.ocr_template_ready = False
        ocr_template_folder = ocr_template_folder or self.config.get("OCR_TEMPLATE_FOLDER")
        statement_folder = statement_folder or self.config.get("STATEMENT_FOLDER")
        statement_template_folder = statement_template_folder or self.config.get("STATEMENT_TEMPLATE_FOLDER")
        self.statement_mode = False
        self.statement_match = None
        self.statement_company = ""
        self.statement_destination_folder = ""
        self.statement_root = os.path.abspath(statement_folder)
        self.statement_template_folder = statement_template_folder
        self.statement_ocr_mode = bool(ocr_mode)

        try:
            if self.ocr_mode:
                os.makedirs(ocr_template_folder, exist_ok=True)
                self.ocr_match = self.check_ocr_templates(filepath, ocr_template_folder)
                self.ocr_template_ready = self.ocr_match is not None
                if self.ocr_match:
                    pct = self.ocr_match["score"] * 100
                    self.log(
                        f"OCR template matched {os.path.basename(self.ocr_match['template']['source'])} "
                        f"at {pct:.1f}% confidence. Human confirmation required.",
                        tag="yellow", display=True,
                    )
                    initial_text = self.ocr_match.get("suggested_filename", "")
                else:
                    self.log(
                        f"No OCR template matched {filename}; manual filename entry is available, "
                        "or you can draw rectangles to train one.",
                        tag="yellow", display=True,
                    )
                    initial_text = ""

                rectangulator, text_box = self.open_rectangulator(
                    filepath, ocr_template_folder, root,
                    scanner=(testing == "scanner"),
                    ocr_mode=True,
                    initial_text=initial_text,
                    ocr_match=self.ocr_match,
                    statement_folder=statement_folder,
                    statement_template_folder=statement_template_folder,
                )
            else:
                self.log(f"Template required for {filename}", display=True)
                if not testing:
                    self.send_email("Must create template", root)
                initial_text = self.suggest_invoice_filename(filepath, ocr=False)
                rectangulator, text_box = self.open_rectangulator(
                    filepath, template_folder, root,
                    scanner=(testing == "scanner"),
                    ocr_mode=False,
                    initial_text=initial_text,
                    statement_folder=statement_folder,
                    statement_template_folder=statement_template_folder,
                )

            if testing == "test":
                return ["test_email"]
            if not rectangulator and not text_box:
                return []
            if not self.invoice:
                self.invoice = True
                entered = self.sanitize_filename(text_box.text or os.path.splitext(filename)[0])
                return ["not_invoice", [
                    os.path.join(os.path.dirname(filepath), f"{entered}.pdf"),
                    self.should_print,
                    self.should_save,
                ]]

            if self.statement_mode:
                target = self.statement_destination(text_box.text)
                if target:
                    return ["statement", target, self.should_print]
                return [None, False]

            # OCR is never authoritative: only the user-approved text box value
            # can become the filename, even when a high-confidence template matched.
            if self.ocr_mode:
                entered = self.sanitize_filename(text_box.text)
                if entered:
                    return [os.path.join(os.path.dirname(filepath), f"{entered}.pdf"), self.should_print]
                return [None, False]

            renamed = rectangulator.rename_pdf()
            if renamed:
                return [renamed, self.should_print]
            entered = self.sanitize_filename(text_box.text)
            if entered:
                return [os.path.join(os.path.dirname(filepath), f"{entered}.pdf"), self.should_print]
        except Exception as e:
            self.log(f"An error occurred while rectangulating: {str(e)}\n{traceback.format_exc()}", tag="red", display=True)
        return [None, False]

    def check_templates(self, pdf_path, template_folder, root):
        # Native templates remain fully automatic and are only used when a real
        # embedded text layer exists. Scans are handled by the separate OCR path.
        if not self.document_has_embedded_text(pdf_path):
            return None

        for template in self._load_native_templates(template_folder):
            try:
                fields = template["fields"]
                vendor_rect, vendor_page = self._field_rect(fields["vendor"], pdf_path)
                identifier = self.sanitize_filename(
                    self.get_text_in_rect(vendor_rect, pdf_path, vendor_page, ocr=False)
                )
                expected = self.sanitize_filename(fields["vendor"].get("expected") or template.get("vendor", ""))
                if not identifier.strip() or self._normalise_match_text(identifier) != self._normalise_match_text(expected):
                    continue

                split_vendors = [v.strip() for v in self.config.get("SPLIT_VENDORS", "").split(",") if v.strip()]
                if expected in split_vendors:
                    with fitz.open(pdf_path) as doc:
                        if len(doc) > 1:
                            self.log(f"Split vendor '{expected}' detected with {len(doc)} pages.")
                            return ["SPLIT_PDF", expected]

                date_rect, date_page = self._field_rect(fields["date"], pdf_path)
                num_rect, num_page = self._field_rect(fields["invoice_number"], pdf_path)
                invoice_date = self.get_text_in_rect(date_rect, pdf_path, date_page, ocr=False)
                invoice_num = self.get_text_in_rect(num_rect, pdf_path, num_page, ocr=False)
                filename = self.build_filename_from_fields(expected, invoice_date, invoice_num)
                if not filename or filename.startswith("_"):
                    continue
                self.log(f"Used template {template['source']} for {expected}")
                return [os.path.join(os.path.dirname(pdf_path), f"{filename}.pdf")]
            except Exception as exc:
                self.log(f"Error checking template {template.get('source', 'unknown')}: {exc}", tag="red", display=True)
        return None

    def open_rectangulator(self, pdf_path, template_folder, root, scanner=False, ocr_mode=False, initial_text="", ocr_match=None, statement_folder=None, statement_template_folder=None):  # Setup the page for the Rectangulator and return the Rectangulator and textbox
        # Reset flags
        self.should_print = True 
        self.should_save = True
        self.hit_submit = False
        self.invoice = True
        self.current_page = 0
        self.ocr_mode = bool(ocr_mode)
        self.ocr_match = ocr_match
        self.statement_mode = False
        self.statement_match = None
        self.statement_company = ""
        self.statement_destination_folder = ""
        self.statement_root = os.path.abspath(statement_folder or self.config.get("STATEMENT_FOLDER", ""))
        self.statement_template_folder = statement_template_folder or self.config.get("STATEMENT_TEMPLATE_FOLDER", "")
        self.statement_ocr_mode = bool(ocr_mode)
        self.fig.patch.set_facecolor(self.palette["surface"])
        self.ax.set_facecolor(self.palette["surface"])

        # Don't print by default when using scanner
        if scanner:
            self.should_print = False  

        self.done_var.set(0)  # reset done variable

        self.rectangulator_instance = None
        self.pdf_path = pdf_path
        self.page_image = None
        self.page_rect = None
        self._render_job = None
        doc_temp = fitz.open(pdf_path)
        self.total_pages = len(doc_temp)
        doc_temp.close()

        def draw_page(page_num):
            if page_num < 0 or page_num >= self.total_pages:
                return

            # Clear rectangles BEFORE wiping the axes, otherwise reset_rectangles
            # tries to remove patches matplotlib has already discarded.
            if self.rectangulator_instance:
                self.rectangulator_instance.reset_rectangles()
                self.rectangulator_instance.page_num = page_num

            doc = fitz.open(pdf_path)
            page = doc[page_num]
            self.page_rect = fitz.Rect(page.rect)
            doc.close()

            self.current_page = page_num
            self.ax.clear()
            self.page_image = None
            self.ax.axis("off")

            # Full page view. Data coordinates are PDF points, not image pixels.
            self.render_view(xlim=(self.page_rect.x0, self.page_rect.x1),
                             ylim=(self.page_rect.y1, self.page_rect.y0))

        self.draw_page = draw_page
        draw_page(self.current_page)


        # Bottom review controls -------------------------------------------------
        # Row 1: filename + actions.  Row 2: optional prefix/postfix training
        # values. Statements are a separate document type but use the same review
        # surface so the operator never has to leave the PDF.
        not_inv_button_ax = self.fig.add_axes([0.71, 0.005, 0.13, 0.07])
        not_inv_button = Button(
            not_inv_button_ax, "Not Invoice",
            color=self.palette["danger_soft"], hovercolor="#F5D5D5")
        not_inv_button.label.set_color(self.palette["danger"])
        not_inv_button.label.set_fontweight("semibold")

        def not_invoice(event):
            try:
                self.invoice = False
                self.ax.clear()
                self.ax.axis("off")
                self.fig.canvas.draw_idle()
                self.done_var.set(1)
            except Exception as e:
                self.log(f"Error in not_invoice: {e}\n{traceback.format_exc()}")
                self.done_var.set(1)

        not_inv_button.on_clicked(not_invoice)

        statement_button_ax = self.fig.add_axes([0.57, 0.005, 0.13, 0.07])
        statement_button = Button(
            statement_button_ax, "Statements",
            color=self.palette["warning_soft"], hovercolor="#F8E2B8")
        statement_button.label.set_color(self.palette["warning"])
        statement_button.label.set_fontweight("semibold")

        # Print / Save checkboxes
        print_checkbox_ax = self.fig.add_axes([0.87, 0.025, 0.025, 0.025])
        print_checkbox_ax.set_facecolor(self.palette["surface_alt"])
        print_checkbox = CheckButtons(print_checkbox_ax, [""], [self.should_print])
        def print_callback(label):
            self.should_print = not self.should_print
        print_checkbox.on_clicked(print_callback)

        save_checkbox_ax = self.fig.add_axes([0.935, 0.025, 0.025, 0.025])
        save_checkbox_ax.set_facecolor(self.palette["surface_alt"])
        save_checkbox = CheckButtons(save_checkbox_ax, [""], [True])
        def save_callback(label):
            self.should_save = not self.should_save
        save_checkbox.on_clicked(save_callback)

        # Preserve the old centering tweak only on Matplotlib versions that expose
        # these implementation details.
        for check in (print_checkbox, save_checkbox):
            if hasattr(check, "lines") and hasattr(check, "rectangles"):
                for i, line in enumerate(check.lines):
                    rect = check.rectangles[i]
                    rect.set_width(1)
                    rect.set_height(1)
                    rect.set_edgecolor("none")
                    center_x = rect.get_width() / 2
                    center_y = rect.get_height() / 2
                    line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
                    line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
                    line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
                    line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])

        print_label = self.fig.text(
            0.852, 0.066, "Print", fontsize=8, color=self.palette["muted"],
            fontweight="semibold")
        save_label = self.fig.text(
            0.918, 0.066, "Save", fontsize=8, color=self.palette["muted"],
            fontweight="semibold")

        # Filename text box
        text_box_ax = self.fig.add_axes([0.08, 0.005, 0.34, 0.07])
        text_box = TextBox(
            text_box_ax, label="", initial=initial_text or "",
            color="#FBFCFC", hovercolor="#FFFFFF")
        for spine in text_box_ax.spines.values():
            spine.set_color(self.palette["border"])
        text_box.text_disp.set_color(self.palette["text"])
        try:
            text_box.set_active(True)
        except Exception as e:
            self.log(f"Error activating text box: {e}\n{traceback.format_exc()}")
        text_box.text_disp.set_horizontalalignment("right")
        text_box.text_disp.set_position((0.95, 0.5))
        text_box.text_disp.set_clip_on(True)
        text_box.text_disp.set_clip_box(text_box_ax.bbox)

        # Prefix / postfix fields are training conveniences. They are saved into
        # PREFIX_VENDORS / POSTFIX_VENDORS when a new invoice template is created.
        known_vendor = ""
        if ocr_match:
            known_vendor = (
                ocr_match.get("template", {}).get("fields", {}).get("vendor", {}).get("expected")
                or ocr_match.get("template", {}).get("vendor", "")
            )
        prefix_initial = self.get_vendor_prefix(known_vendor) if known_vendor else ""
        postfix_initial = self.get_vendor_postfix(known_vendor) if known_vendor else ""
        prefix_ax = self.fig.add_axes([0.58, 0.105, 0.16, 0.045])
        postfix_ax = self.fig.add_axes([0.78, 0.105, 0.16, 0.045])
        prefix_box = TextBox(prefix_ax, label="", initial=prefix_initial, color="#FBFCFC", hovercolor="#FFFFFF")
        postfix_box = TextBox(postfix_ax, label="", initial=postfix_initial, color="#FBFCFC", hovercolor="#FFFFFF")
        for control_ax, control in ((prefix_ax, prefix_box), (postfix_ax, postfix_box)):
            for spine in control_ax.spines.values():
                spine.set_color(self.palette["border"])
            control.text_disp.set_color(self.palette["text"])
        prefix_label = self.fig.text(0.58, 0.153, "Prefix (template)", fontsize=8, color=self.palette["muted"], fontweight="semibold")
        postfix_label = self.fig.text(0.78, 0.153, "Postfix (template)", fontsize=8, color=self.palette["muted"], fontweight="semibold")

        text_box_label = self.fig.text(
            0.08, 0.082, "Filename - Invoice Date _ Invoice #",
            fontsize=8.5, color=self.palette["muted"], fontweight="semibold")

        if self.ocr_mode and ocr_match:
            confidence = ocr_match.get("score", 0.0) * 100
            line1 = f"- OCR template match: {confidence:.1f}% — review the suggested filename"
            line2 = "- Edit or paste a filename and Submit; drawing rectangles is optional"
            line3 = "- To retrain: blue Vendor → amber Date → purple Invoice #; right-click to verify"
        elif self.ocr_mode:
            line1 = "- OCR mode: manual filename entry is allowed without creating a template"
            line2 = "- Optional training: blue Vendor → amber Date → purple Invoice #, then right-click"
            line3 = "- Prefix / postfix above are saved when you create an invoice template"
        else:
            line1 = "- Draw 3 boxes: blue Vendor → amber Date → purple Invoice #"
            line2 = "- Or type/paste the filename manually if the PDF does not contain every field"
            line3 = "- Right-click verify  •  Middle-drag pan  •  Scroll zoom"
        instruction_color = self.palette["warning"] if self.ocr_mode else self.palette["pewter_dark"]
        instruction_label = self.fig.text(
            0.08, 0.975, line1, fontsize=9.5, color=instruction_color, fontweight="semibold")
        instruction_label_2 = self.fig.text(
            0.08, 0.95, line2, fontsize=9, color=self.palette["text"])
        instruction_label_3 = self.fig.text(
            0.08, 0.925, line3, fontsize=9, color=self.palette["muted"])

        def set_statement_instructions(match=None):
            instruction_label.set_text("- STATEMENT mode: blue Identifier → amber Date (optional template training)")
            if match:
                pct = match.get("score", 0.0) * 100
                instruction_label_2.set_text(
                    f"- Statement template matched {match.get('company', '')} at {pct:.1f}% — review the filename")
            else:
                instruction_label_2.set_text(
                    "- Folder selected from Statements; type/paste a filename or draw Identifier + Date and right-click")
            instruction_label_3.set_text("- Statements always save; scanner/MFP statements are not printed")
            instruction_label.set_color(self.palette["warning"])
            text_box_label.set_text("Statement filename  •  Company Statement Date")
            self.fig.canvas.draw_idle()

        def choose_statement(event=None):
            try:
                os.makedirs(self.statement_root, exist_ok=True)
                use_ocr = bool(self.ocr_mode or not self.document_has_embedded_text(pdf_path))
                match = self.check_statement_templates(
                    pdf_path, self.statement_template_folder, self.statement_root, ocr=use_ocr)
                selected_folder = ""
                company = ""
                if match:
                    selected_folder = match.get("destination_folder", "")
                    company = match.get("company", "")
                    if selected_folder:
                        os.makedirs(selected_folder, exist_ok=True)
                if not selected_folder:
                    selected_folder = filedialog.askdirectory(
                        parent=root.root,
                        title="Choose the company folder for this statement",
                        initialdir=self.statement_root,
                        mustexist=True,
                    )
                    if not selected_folder:
                        return
                    root_abs = os.path.abspath(self.statement_root)
                    selected_folder = os.path.abspath(selected_folder)
                    try:
                        if os.path.commonpath([root_abs, selected_folder]) != root_abs:
                            self.create_alert("Please choose a company folder inside the configured Statements folder.")
                            return
                    except ValueError:
                        self.create_alert("Please choose a company folder inside the configured Statements folder.")
                        return
                    if selected_folder == root_abs:
                        self.create_alert("Please choose the company's folder, not the Statements root itself.")
                        return
                    company = os.path.basename(selected_folder.rstrip("\\/"))

                self.statement_mode = True
                self.statement_match = match
                self.statement_company = self.sanitize_filename(company)
                self.statement_destination_folder = selected_folder
                self.statement_ocr_mode = use_ocr

                # Statements are always saved. They are printed for ordinary email
                # intake, but not when the source is the configured scanner/MFP.
                desired_print = not scanner
                try:
                    if bool(print_checkbox.get_status()[0]) != desired_print:
                        print_checkbox.set_active(0)
                    if not bool(save_checkbox.get_status()[0]):
                        save_checkbox.set_active(0)
                except Exception:
                    pass
                self.should_print = desired_print
                self.should_save = True
                if self.rectangulator_instance is not None:
                    self.rectangulator_instance.statement_mode = True
                    self.rectangulator_instance.statement_ocr_mode = use_ocr
                    self.rectangulator_instance.reset_rectangles()

                if match and match.get("suggested_filename"):
                    text_box.set_val(match["suggested_filename"])
                    self.log(
                        f"Statement template selected {self.statement_company}; review the date/filename and Submit.",
                        tag="yellow", display=True)
                else:
                    text_box.set_val(f"{self.statement_company} Statement ")
                    self.log(
                        f"Statement folder selected: {selected_folder}. You may enter the date manually or train a template.",
                        tag="yellow", display=True)
                set_statement_instructions(match)
            except Exception as exc:
                self.log(f"Unable to start Statement mode: {exc}\n{traceback.format_exc()}", tag="red", display=True)

        statement_button.on_clicked(choose_statement)

        def on_text_submit(event=None):
            if self.hit_submit:
                return
            self.hit_submit = True
            try:
                entered = self.sanitize_filename(text_box.text)
                if not entered:
                    self.create_alert("Please enter a filename before submitting.")
                    return
                # OCR templates are optional. A manually entered filename is a
                # complete, valid path through the workflow.
                if self.statement_mode and not self.statement_destination_folder:
                    self.create_alert("Choose the Statement company folder first.")
                    return
                if entered != text_box.text:
                    text_box.set_val(entered)
                filename_is_correct = self.create_alert(f"Is '{entered}' the correct filename?")
                if filename_is_correct:
                    self.ax.clear()
                    self.ax.axis("off")
                    self.fig.canvas.draw_idle()
                    self.done_var.set(1)
            finally:
                self.hit_submit = False

        submit_button_ax = self.fig.add_axes([0.43, 0.005, 0.13, 0.07])
        submit_button = Button(
            submit_button_ax, "Submit", color=self.palette["accent"],
            hovercolor=self.palette["accent_hover"])
        submit_button.label.set_color("#FFFFFF")
        submit_button.label.set_fontweight("semibold")
        submit_button.on_clicked(on_text_submit)

        # Matplotlib's TextBox does not consistently receive Ctrl+V on TkAgg.
        # Bind the underlying Tk canvas so clipboard paste works reliably.
        canvas_widget = self.fig.canvas.get_tk_widget()
        paste_bindings = []
        def paste_filename(_event=None):
            try:
                clip = root.root.clipboard_get()
            except (tk.TclError, AttributeError):
                return "break"
            text_box.set_val(f"{text_box.text}{clip}")
            return "break"
        for sequence in ("<Control-v>", "<Control-V>", "<Shift-Insert>"):
            try:
                bind_id = canvas_widget.bind(sequence, paste_filename, add="+")
                paste_bindings.append((sequence, bind_id))
            except tk.TclError:
                pass

        # Page navigation
        prev_button_ax = self.fig.add_axes([0.02, 0.45, 0.04, 0.1])
        prev_button = Button(
            prev_button_ax, "‹", color=self.palette["surface_alt"], hovercolor="#E8EDF0")
        prev_button.label.set_color(self.palette["pewter_dark"])
        prev_button.label.set_fontsize(14)
        def on_prev(event):
            if self.current_page > 0:
                self.current_page -= 1
                draw_page(self.current_page)
        prev_button.on_clicked(on_prev)

        next_button_ax = self.fig.add_axes([0.94, 0.45, 0.04, 0.1])
        next_button = Button(
            next_button_ax, "›", color=self.palette["surface_alt"], hovercolor="#E8EDF0")
        next_button.label.set_color(self.palette["pewter_dark"])
        next_button.label.set_fontsize(14)
        def on_next(event):
            if self.current_page < self.total_pages - 1:
                self.current_page += 1
                draw_page(self.current_page)
        next_button.on_clicked(on_next)

        self.fig.canvas.draw_idle()
        rectangulator = Rectangulator(
            self.ax, self.fig, pdf_path, template_folder, self,
            text_box=text_box, ocr_mode=self.ocr_mode,
            prefix_box=prefix_box, postfix_box=postfix_box,
            statement_template_folder=self.statement_template_folder,
            statement_root=self.statement_root,
        )
        self.rectangulator_instance = rectangulator
        rectangulator.page_num = self.current_page
        root.root.wait_variable(self.done_var)

        # Cleanup controls and callbacks before the next queued document.
        for sequence, bind_id in paste_bindings:
            try:
                canvas_widget.unbind(sequence, bind_id)
            except (tk.TclError, TypeError):
                pass
        for control_ax in [
            text_box_ax, prefix_ax, postfix_ax, not_inv_button_ax, statement_button_ax,
            print_checkbox_ax, save_checkbox_ax, submit_button_ax, prev_button_ax, next_button_ax,
        ]:
            try:
                self.fig.delaxes(control_ax)
            except (KeyError, ValueError):
                pass
        for label in [
            text_box_label, prefix_label, postfix_label, instruction_label,
            instruction_label_2, instruction_label_3, print_label, save_label,
        ]:
            try:
                label.remove()
            except ValueError:
                pass
        for cid in rectangulator.cids:
            self.fig.canvas.mpl_disconnect(cid)
        rectangulator.cids.clear()
        self.ax.clear()
        self.ax.axis("off")
        self.fig.canvas.draw_idle()

        return rectangulator, text_box

    def render_view(self, xlim=None, ylim=None):
        # Re-render only the visible slice of the page, at the resolution the
        # canvas can actually display. The image is placed with an extent
        # measured in PDF points, so axes data coordinates always match
        # page.get_text("words") output -- existing templates keep working.
        if self.page_rect is None or not self.pdf_path:
            return
        doc = None
        try:
            if xlim is None:
                xlim = self.ax.get_xlim()
            if ylim is None:
                ylim = self.ax.get_ylim()
            x0, x1 = sorted(xlim)
            y0, y1 = sorted(ylim)

            clip = fitz.Rect(x0, y0, x1, y1) & self.page_rect
            if clip.is_empty or clip.width <= 0 or clip.height <= 0:
                return

            # Points-to-screen-pixels scale for the current view
            bbox = self.ax.get_window_extent()
            scale = min(max(bbox.width, 1.0) / clip.width,
                        max(bbox.height, 1.0) / clip.height)
            zoom = max(1.0, min(scale * self.SUPERSAMPLE, self.MAX_RENDER_ZOOM))

            # Keep the bitmap within a sane memory budget
            pixels = (clip.width * zoom) * (clip.height * zoom)
            if pixels > self.MAX_RENDER_PIXELS:
                zoom *= (self.MAX_RENDER_PIXELS / pixels) ** 0.5

            doc = fitz.open(self.pdf_path)
            page = doc[min(max(self.current_page, 0), len(doc) - 1)]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

            # pix.irect is in the zoomed space; divide back into PDF points.
            # Indexed rather than .x0/.y0 -- some PyMuPDF versions return a tuple.
            ir = tuple(pix.irect)
            extent = [ir[0] / zoom, ir[2] / zoom, ir[3] / zoom, ir[1] / zoom]

            if self.page_image is None:
                self.page_image = self.ax.imshow(img_array, extent=extent)
                self.ax.axis("off")
            else:
                self.page_image.set_data(img_array)
                self.page_image.set_extent(extent)
                # set_extent can rescale the axes; restore the user's view
                self.ax.set_xlim(xlim)
                self.ax.set_ylim(ylim)
            self.fig.canvas.draw_idle()
        except Exception as e:
            self.log(f"Error rendering page view: {str(e)} \n{traceback.format_exc()}")
        finally:
            if doc:
                doc.close()

    def schedule_render(self, delay=120):
        # Coalesce bursts of scroll / pan / resize events into one re-render
        try:
            tk_root = self.root.root
            if self._render_job is not None:
                tk_root.after_cancel(self._render_job)
            self._render_job = tk_root.after(delay, self._run_scheduled_render)
        except Exception:
            self.render_view()

    def _run_scheduled_render(self):
        self._render_job = None
        self.render_view()

    def print_invoice(self, filepath):
        # Printing is centralized in EmailProcessor so both normal and manual
        # paths honor PRINTER_NAME consistently.
        if self.root and hasattr(self.root, "print_invoice"):
            return self.root.print_invoice(filepath)
        return False

    def log(self, *args, tag="purple", send_email=False, display=False):
        message = " ".join(str(arg) for arg in args)
        if self.root and hasattr(self.root, "log"):
            # Root owns disk logging / alert-email dispatch, preventing duplicate
            # lines and duplicate alert emails from Rectangulator.
            self.root.log(message, tag=tag, send_email=send_email, console=True)
        else:
            print(f"-{message}")

    def sanitize_filename(self, filename):  # Remove invalid characters from the filename
        sanitized_filename = re.sub(r"[^\w_. -]", "", filename.replace("/", "-"))
        return sanitized_filename.strip()

    def get_text_in_rect(self, rect, pdf_path, page_num=0, ocr=False):
        try:
            selection = self._normalise_selection(rect)
            if selection.is_empty or selection.width <= 0 or selection.height <= 0:
                return ""
            selected = []
            for word in self._get_page_words(pdf_path, page_num, ocr=ocr):
                word_rect = fitz.Rect(*map(float, word[:4]))
                if word_rect.is_empty:
                    continue
                center = fitz.Point((word_rect.x0 + word_rect.x1) / 2, (word_rect.y0 + word_rect.y1) / 2)
                intersection = word_rect & selection
                overlap = 0.0 if intersection.is_empty else (intersection.get_area() / max(word_rect.get_area(), 1e-9))
                if selection.contains(center) or overlap >= 0.30:
                    selected.append(word)
            # PyMuPDF already returns sort=True visual order from the cache.
            return " ".join(str(word[4]) for word in selected).strip()
        except Exception as e:
            self.log(f"An error occurred while processing the PDF: {e} {traceback.format_exc()}", tag="red")
            return ""

    def get_vendor_rule(self, setting_key, identifier):
        raw = str(self.config.get(setting_key, "") or "")
        for pair in [p.strip() for p in raw.split(",") if ":" in p]:
            vendor, value = pair.split(":", 1)
            if vendor.strip().casefold() == str(identifier or "").strip().casefold():
                return value.strip()
        return ""

    def get_vendor_prefix(self, identifier):
        return self.get_vendor_rule("PREFIX_VENDORS", identifier)

    def get_vendor_postfix(self, identifier):
        return self.get_vendor_rule("POSTFIX_VENDORS", identifier)

    def check_date_outlier(self, invoice_name, invoice_date):  # Check if the date is an outlier and correct it
        calendar = {
            "Jan": "01",
            "Feb": "02",
            "Mar": "03",
            "Apr": "04",
            "May": "05",
            "Jun": "06",
            "Jul": "07",
            "Aug": "08",
            "Sep": "09",
            "Oct": "10",
            "Nov": "11",
            "Dec": "12"
        }
        calendar2 = {
            "January": "01",
            "February": "02",
            "March": "03",
            "April": "04",
            "May": "05",
            "June": "06",
            "July": "07",
            "August": "08",
            "September": "09",
            "October": "10",
            "November": "11",
            "December": "12"
        }

        try:
            # Uses format "DD-Month-YY"
            invoice_date_copy = invoice_date.split("-")
            day = invoice_date_copy[0]
            invoice_date_copy[0] = calendar[invoice_date_copy[1]]
            invoice_date_copy[1] = day
            invoice_date_copy = "-".join(invoice_date_copy)
            if datetime.strptime(invoice_date_copy, "%m-%d-%y"):
                return invoice_date_copy
        except:
            pass
        
        try:
            # Uses format "Month DD, YYYY"
            invoice_date_copy = invoice_date.replace(",", "").split(" ")
            invoice_date_copy[0] = calendar2[invoice_date_copy[0]]
            invoice_date_copy = "/".join(invoice_date_copy)
            if datetime.strptime(invoice_date_copy, "%m/%d/%Y"):
                return invoice_date_copy
        except:
            pass

        try:
            # Uses format "Mon DD, YYYY"
            invoice_date_copy = invoice_date.replace(",", "").split(" ")
            invoice_date_copy[0] = calendar[invoice_date_copy[0]]
            invoice_date_copy = "/".join(invoice_date_copy)
            if datetime.strptime(invoice_date_copy, "%m/%d/%Y"):
                return invoice_date_copy
        except:
            pass

        return self.clean_date(invoice_date.strip())

    def clean_date(self, invoice_date):  # Clean the date to be in the format "MM-DD-YY"
        raw = " ".join(str(invoice_date or "").strip().split())
        date_patterns = [
            "%B %d, %Y", "%b %d, %Y", "%B %d, %y", "%b %d, %y",
            "%B %d %Y", "%b %d %Y", "%B %d %y", "%b %d %y",
            "%d-%B-%Y", "%d-%b-%Y", "%d-%B-%y", "%d-%b-%y",
            "%d %B %Y", "%d %b %Y", "%d %B %y", "%d %b %y",
            "%m-%d-%Y", "%m-%d-%y", "%m/%d/%Y", "%m/%d/%y",
            "%Y-%m-%d", "%y-%m-%d", "%Y/%m/%d", "%y/%m/%d",
        ]

        candidates = [raw]
        # Rectangles sometimes include labels such as "Statement Date:" or
        # "Invoice Date:". Pull likely date substrings out before giving up.
        date_regexes = [
            r"(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+\d{1,2},?\s+\d{2,4}",
            r"\d{1,2}[- ](?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[- ,]+\d{2,4}",
            r"\d{1,4}[/-]\d{1,2}[/-]\d{1,4}",
        ]
        for pattern in date_regexes:
            match = re.search(pattern, raw, re.I)
            if match:
                candidate = match.group(0).replace("Sept ", "Sep ").replace("Sept-", "Sep-")
                if candidate not in candidates:
                    candidates.append(candidate)

        for candidate in candidates:
            for pattern in date_patterns:
                try:
                    dt = datetime.strptime(candidate, pattern)
                    return dt.strftime("%m-%d-%y")
                except ValueError:
                    continue
        self.log(f"Could not convert {raw} to date")
        return raw

    def send_email(self, body, root):  # Sends email to me
        if root:
            sender_email = f"{root.username}.sndex@gmail.com"
            password = root.password
            if root.TESTING:
                return
        else:
            sender_email = f"{self.config['APC_USER']}.sndex@gmail.com"
            password = get_password(self.config['APC_USER'])

        try:
            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = self.config["RECEIVER_EMAIL"]
            message.attach(MIMEText(body, "plain"))

            # Send the email
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, password)
                server.sendmail(sender_email, self.config["RECEIVER_EMAIL"], message.as_string())
            self.log(f"Email sent successfully: {body}")
        except Exception as e:
            self.log(f"Error sending email {body} - {str(e)}")

    def create_alert(self, message, numbered_buttons=0):  # Create an alert window for user input
        try:
            parent = self.root.alert_container
            panel = AlertWindow(parent, message, numbered_buttons)
            panel.pack(fill=tk.BOTH, expand=True)
            parent.lift()
            panel.grab_set()
            panel.focus_set()
            answer = panel.get_answer()  # wait for user input
            panel.destroy()  # destroy the alert window
            parent.lower()  # lower the alert container
            return answer
        except Exception as e:
            self.log(f"Error creating alert: {str(e)} \n{traceback.format_exc()}")
            return False

class Rectangulator:

    def __init__(self, ax, fig, pdf_path, template_folder, rectangulator_handler, text_box=None, ocr_mode=False,
                 prefix_box=None, postfix_box=None, statement_template_folder=None, statement_root=None):
        self.rectangulator_handler = rectangulator_handler
        self.pdf_path = pdf_path
        self.template_folder = template_folder
        self.fig = fig
        self.ax = ax
        self.text_box = text_box
        self.prefix_box = prefix_box
        self.postfix_box = postfix_box
        self.ocr_mode = bool(ocr_mode)
        self.statement_mode = False
        self.statement_ocr_mode = False
        self.statement_template_folder = statement_template_folder or rectangulator_handler.config.get("STATEMENT_TEMPLATE_FOLDER", "")
        self.statement_root = statement_root or rectangulator_handler.config.get("STATEMENT_FOLDER", "")

        self.rectangles = []  # contains rectangle objects
        self.coordinates = []  # contains coordinates of rectangle objects
        self.correcting_rect_index = None  # used when redrawing specific rectangle
        self.start_x = None
        self.start_y = None
        self.rect = None
        self.zoom_factor = 1.2
        self.pan_factor = 1
        self.pan_start = None
        self.prev_x = None
        self.prev_y = None
        self.initial_xlim = self.ax.get_xlim()
        self.initial_ylim = self.ax.get_ylim()

        # Connect the event handlers to the canvas
        self.cids = []
        cid1 = self.ax.figure.canvas.mpl_connect("button_press_event", self.on_button_press)
        cid2 = self.ax.figure.canvas.mpl_connect("button_release_event", self.on_button_release)
        cid3 = self.ax.figure.canvas.mpl_connect("motion_notify_event", self.on_move)
        cid4 = self.ax.figure.canvas.mpl_connect("scroll_event", self.on_scroll)
        cid5 = self.ax.figure.canvas.mpl_connect("key_press_event", self.on_key_press)
        cid6 = self.ax.figure.canvas.mpl_connect("resize_event", self.on_resize)
        self.cids.extend([cid1, cid2, cid3, cid4, cid5, cid6])

    def rename_pdf(self):
        try:
            extracted_texts = [
                self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path, self.page_num, ocr=False)
                for rect in self.rectangles
            ]
            if len(self.rectangles) == 3 and all(extracted_texts):
                self.rectangulator_handler.log("Creating new native-text template")
                self.save_template(ocr=False)
                filename = self.rectangulator_handler.build_filename_from_fields(
                    extracted_texts[0], extracted_texts[1], extracted_texts[2]
                )
                return os.path.join(os.path.dirname(self.pdf_path), f"{filename}.pdf") if filename else None
        except RecursionError:
            self.rectangulator_handler.log("Window closed please try again")
        except Exception as e:
            self.rectangulator_handler.log(traceback.format_exc())
            self.rectangulator_handler.log(f"Error occurred while renaming: {e}")
        return None

    def save_template(self, ocr=False):
        if len(self.rectangles) != 3:
            return False
        extracted = [
            self.rectangulator_handler.get_text_in_rect(
                rect, self.pdf_path, self.page_num, ocr=ocr
            ) for rect in self.rectangles
        ]
        if not all(extracted):
            self.rectangulator_handler.log("One or more selected fields were empty; template was not created.", tag="orange")
            return False

        vendor = self.rectangulator_handler.sanitize_filename(extracted[0])
        if not vendor:
            return False

        prefix = str(self.prefix_box.text if self.prefix_box is not None else "").strip()
        postfix = str(self.postfix_box.text if self.postfix_box is not None else "").strip()
        if prefix or postfix:
            root = self.rectangulator_handler.root
            if root and hasattr(root, "update_vendor_affixes"):
                root.update_vendor_affixes(
                    vendor,
                    prefix=prefix if prefix else None,
                    postfix=postfix if postfix else None,
                )
                # refresh_config picks up the persisted rule before rename_pdf
                # constructs this very invoice's filename.
                self.rectangulator_handler.refresh_config()

        os.makedirs(self.template_folder, exist_ok=True)
        safe_vendor = re.sub(r"[^A-Za-z0-9_. -]+", "_", vendor).strip() or "template"
        suffix = "ocr" if ocr else "text"
        filename = os.path.join(self.template_folder, f"{safe_vendor}.{suffix}.json")
        fields = {}
        for name, rect, expected in zip(("vendor", "date", "invoice_number"), self.rectangles, extracted):
            fields[name] = self.rectangulator_handler.normalized_field(
                rect, self.pdf_path, self.page_num, expected
            )
        variant = {"vendor": vendor, "fields": fields}
        payload = {"version": 2, "type": "ocr" if ocr else "text", "vendor": vendor, "templates": []}
        if os.path.exists(filename):
            try:
                with open(filename, "r", encoding="utf-8") as f:
                    old = json.load(f)
                if old.get("type") == payload["type"]:
                    payload = old
                    payload.setdefault("templates", [])
            except Exception:
                pass
        payload["templates"].append(variant)
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        self.rectangulator_handler.invalidate_template_cache(self.template_folder)
        self.rectangulator_handler.log(
            f"Created {'OCR' if ocr else 'native-text'} invoice template {filename}"
        )
        return True

    def save_statement_template(self):
        if len(self.rectangles) != 2:
            return False
        ocr = bool(self.statement_ocr_mode)
        extracted = [
            self.rectangulator_handler.get_text_in_rect(
                rect, self.pdf_path, self.page_num, ocr=ocr
            ) for rect in self.rectangles
        ]
        if not all(extracted):
            self.rectangulator_handler.log(
                "Identifier or statement date was empty; Statement template was not created.",
                tag="orange", display=True)
            return False

        company = self.rectangulator_handler.sanitize_filename(
            self.rectangulator_handler.statement_company)
        destination = os.path.abspath(
            self.rectangulator_handler.statement_destination_folder or "")
        root = os.path.abspath(self.statement_root or "")
        if not company or not destination or not root:
            return False
        try:
            if os.path.commonpath([root, destination]) != root:
                return False
        except ValueError:
            return False

        rel_folder = os.path.relpath(destination, root)
        folder = os.path.abspath(self.statement_template_folder)
        os.makedirs(folder, exist_ok=True)
        safe_company = re.sub(r"[^A-Za-z0-9_. -]+", "_", company).strip() or "statement"
        filename = os.path.join(folder, f"{safe_company}.statement.json")
        fields = {
            "identifier": self.rectangulator_handler.normalized_field(
                self.rectangles[0], self.pdf_path, self.page_num, extracted[0]),
            "date": self.rectangulator_handler.normalized_field(
                self.rectangles[1], self.pdf_path, self.page_num, extracted[1]),
        }
        variant = {
            "company": company,
            "folder": rel_folder,
            "ocr": ocr,
            "fields": fields,
        }
        payload = {"version": 1, "type": "statement", "company": company, "templates": []}
        if os.path.exists(filename):
            try:
                with open(filename, "r", encoding="utf-8") as f:
                    old = json.load(f)
                if old.get("type") == "statement":
                    payload = old
                    payload.setdefault("templates", [])
            except Exception:
                pass
        payload["templates"].append(variant)
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        self.rectangulator_handler.invalidate_template_cache(folder)
        self.rectangulator_handler.log(
            f"Created {'OCR ' if ocr else ''}Statement template {filename}",
            tag="lgreen", display=True)
        return True

    def on_key_press(self, event):  # Handle key press events
        # Matplotlib TextBox does not implement clipboard paste consistently on
        # TkAgg. Handle Ctrl+V here when the filename box is the active typing
        # target. The Tk-level binding in open_rectangulator is retained as an
        # additional Windows fallback.
        if str(getattr(event, "key", "")).lower() in ("ctrl+v", "cmd+v"):
            if self.text_box is not None and getattr(self.text_box, "capturekeystrokes", False):
                try:
                    clip = self.rectangulator_handler.root.root.clipboard_get()
                    self.text_box.set_val(f"{self.text_box.text}{clip}")
                except (tk.TclError, AttributeError):
                    pass
            return

        if event.key == "escape":  # reset zoom and position
            page_rect = getattr(self.rectangulator_handler, "page_rect", None)
            if page_rect is not None:
                self.ax.set_xlim(page_rect.x0, page_rect.x1)
                self.ax.set_ylim(page_rect.y1, page_rect.y0)
            else:
                self.ax.set_xlim(self.initial_xlim)
                self.ax.set_ylim(self.initial_ylim)
            self.pan_start = None
            self.ax.figure.canvas.draw()
            self.rectangulator_handler.schedule_render(delay=0)

    def on_resize(self, event):  # Re-render to match the new canvas resolution
        self.rectangulator_handler.schedule_render(delay=250)

    def on_button_press(
            self, event):  # Handle left and right mouse button press events
        if event.button == 1:  # left mouse button, draw rectangles
            # Ignore if the mouse click is outside the plot area
            if event.xdata is None or event.ydata is None:
                return

            # Start drawing a rectangle. Each required field has a consistent
            # color so the selection order is visible at a glance.
            self.start_x = event.xdata
            self.start_y = event.ydata
            if self.correcting_rect_index is not None:
                field_index = self.correcting_rect_index
            else:
                field_index = min(len(self.rectangles), 2)
            edge_color = self.rectangulator_handler.field_colors[field_index]
            self.rect = Rectangle(
                (self.start_x, self.start_y), 0, 0, edgecolor=edge_color,
                linewidth=2.2, fill=False)
            self.ax.add_patch(self.rect)
            self.ax.figure.canvas.draw()

        elif event.button == 2:  # middle mouse button, pan
            self.pan_start = (event.x, event.y)

        elif event.button == 3:  # right mouse button, verify / save rectangles
            required = 2 if self.statement_mode else 3
            if len(self.rectangles) != required:
                label = "two (Identifier and Date)" if self.statement_mode else "three"
                self.rectangulator_handler.log(
                    f"Please draw exactly {label} rectangles", display=True)
                self.reset_rectangles()
                return
            self.verify_selection()

    def on_button_release(self, event):  # Handle key release events
        if event.button == 1 and self.rect:  # left mouse button, save rectangle
            self.start_x = None
            self.start_y = None

            # Normalize drag direction so left/up drags are valid too.
            normalized = self.rectangulator_handler._normalise_selection(self.rect)
            self.rect.set_x(normalized.x0)
            self.rect.set_y(normalized.y0)
            self.rect.set_width(normalized.width)
            self.rect.set_height(normalized.height)
            self.rectangles.append(self.rect)
            self.coordinates.append((normalized.x0, normalized.y0, normalized.width, normalized.height))
            self.rect = None
            self.ax.figure.canvas.draw()

        elif event.button == 2:  # middle mouse button, stop panning
            self.pan_start = None
            self.prev_x = None
            self.prev_y = None
            self.rectangulator_handler.schedule_render()

    def on_move(self, event):  # Continuously update drawn rectangles
        # Pan if middle mouse button is pressed
        if event.button == 2 and self.pan_start:
            if self.prev_x is not None and self.prev_y is not None:
                xlim = self.ax.get_xlim()
                ylim = self.ax.get_ylim()

                # Convert screen pixels to data units so panning tracks the
                # cursor at any zoom level instead of drifting
                bbox = self.ax.get_window_extent()
                scale = abs(xlim[1] - xlim[0]) / max(bbox.width, 1.0)
                dx = (event.x - self.prev_x) * scale * self.pan_factor
                dy = (event.y - self.prev_y) * scale * self.pan_factor
                new_xlim = xlim[0] - dx, xlim[1] - dx
                new_ylim = ylim[0] + dy, ylim[1] + dy

                self.ax.set_xlim(new_xlim)
                self.ax.set_ylim(new_ylim)
                self.ax.figure.canvas.draw_idle()

            self.prev_x = event.x
            self.prev_y = event.y
            return

        if self.rect is None or self.start_x is None or self.start_y is None:
            return

        current_x = event.xdata
        current_y = event.ydata
        # If outside plot area, set to 0
        if current_x is None:
            current_x = 0
        if current_y is None:
            current_y = 0

        # Get rectangle width and height
        width = current_x - self.start_x
        height = current_y - self.start_y

        # Update rectangle
        self.rect.set_width(width)
        self.rect.set_height(height)
        self.ax.figure.canvas.draw()

    def on_scroll(self, event):  # Zoom in and out
        if event.button == "down":  # out
            self.zoom(event.xdata, event.ydata, 1 / self.zoom_factor)
        elif event.button == "up":  # in
            self.zoom(event.xdata, event.ydata, self.zoom_factor)

    def reset_rectangles(self, specific_rect=None):  # Reset the current rectangles
        if specific_rect is not None:
            self.rectangles[specific_rect].remove()
            self.rectangles.pop(specific_rect)
            self.coordinates.pop(specific_rect)
            self.ax.figure.canvas.draw()
        else:
            for rect in self.rectangles:
                rect.remove()
            self.rectangles = []
            self.coordinates = []
            self.rect = None
            self.correcting_rect_index = None
            self.ax.figure.canvas.draw()

    def zoom(self, x, y, zoom_factor):  # Zoom in and out with scroll wheel
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        if xlim is None or ylim is None:
            return

        # Calculate new limits
        if x is None:
            x = np.mean(xlim)
        if y is None:
            y = np.mean(ylim)
        new_xlim = ((xlim[0] - x) / zoom_factor) + x, (
            (xlim[1] - x) / zoom_factor) + x
        new_ylim = ((ylim[0] - y) / zoom_factor) + y, (
            (ylim[1] - y) / zoom_factor) + y

        self.ax.set_xlim(new_xlim)
        self.ax.set_ylim(new_ylim)
        self.ax.figure.canvas.draw()
        # Re-render the newly visible region at full detail
        self.rectangulator_handler.schedule_render()

    def verify_selection(self):
        if self.correcting_rect_index is not None:
            corrected_rect = self.rectangles.pop()
            corrected_coord = self.coordinates.pop()
            self.rectangles.insert(self.correcting_rect_index, corrected_rect)
            self.coordinates.insert(self.correcting_rect_index, corrected_coord)
            self.correcting_rect_index = None

        if self.statement_mode:
            headers = ["--- Statement Identifier: ", "--- Statement Date: "]
            use_ocr = bool(self.statement_ocr_mode)
        else:
            headers = ["--- Company Name: ", "--- Invoice Date: ", "--- Invoice Number: "]
            use_ocr = bool(self.ocr_mode)

        extracted_values = [
            self.rectangulator_handler.get_text_in_rect(
                rect, self.pdf_path, self.page_num, ocr=use_ocr
            ) for rect in self.rectangles
        ]
        extracted_text = "\n".join(h + v for h, v in zip(headers, extracted_values))
        text_is_correct = self.rectangulator_handler.create_alert(
            f"Does the following text match what you selected?\n\n{extracted_text}",
            numbered_buttons=len(headers),
        )

        if isinstance(text_is_correct, int) and not isinstance(text_is_correct, bool):
            if 0 <= text_is_correct < len(headers):
                self.correcting_rect_index = text_is_correct
                self.rectangulator_handler.log(
                    f"Please reselect {headers[self.correcting_rect_index]}")
                self.reset_rectangles(specific_rect=self.correcting_rect_index)
            return

        if not text_is_correct:
            self.rectangulator_handler.log("Please reselect rectangles", display=True)
            self.reset_rectangles()
            return

        if self.statement_mode:
            if not all(extracted_values):
                self.rectangulator_handler.log(
                    "The identifier and date must both be readable to create a Statement template. "
                    "You can still type the filename manually and Submit.",
                    tag="orange", display=True)
                return
            if self.save_statement_template():
                suggestion = self.rectangulator_handler.build_statement_filename(
                    self.rectangulator_handler.statement_company, extracted_values[1])
                if self.text_box is not None:
                    self.text_box.set_val(suggestion)
                self.rectangulator_handler.log(
                    f"Statement template trained. Suggested filename: '{suggestion}'. Review/edit and Submit.",
                    tag="yellow", display=True)
                self.reset_rectangles()
            return

        if self.ocr_mode:
            if not all(extracted_values):
                self.rectangulator_handler.log(
                    "OCR did not read all three selections. Redraw the empty field, or simply enter the filename manually.",
                    tag="orange", display=True)
                return
            if self.save_template(ocr=True):
                suggestion = self.rectangulator_handler.build_filename_from_fields(*extracted_values)
                self.rectangulator_handler.ocr_template_ready = True
                if self.text_box is not None:
                    self.text_box.set_val(suggestion)
                self.rectangulator_handler.log(
                    f"OCR suggested '{suggestion}'. Review/edit the filename and click Submit.",
                    tag="yellow", display=True)
                self.reset_rectangles()
            return

        # Native invoice template: right-click confirmation keeps the original
        # behavior of immediately completing the review.
        self.ax.clear()
        self.ax.axis("off")
        self.fig.canvas.draw_idle()
        self.rectangulator_handler.done_var.set(1)

