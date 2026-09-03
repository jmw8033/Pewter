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
config.setdefault("TEST_OCR_TEMPLATE_FOLDER", os.path.join(_test_template_parent, "OCR_Templates"))
config.setdefault("OCR_FUZZY_THRESHOLD", 0.72)
config.setdefault("OCR_DPI", 250)
config.setdefault("OCR_LANGUAGE", "eng")
config.setdefault("TESSDATA_PREFIX", "")
config.setdefault("PRINTER_NAME", "")
config.setdefault("MIN_EMBEDDED_TEXT_CHARS", 40)

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

        # Small LRU caches: native text and OCR are both expensive enough that
        # reopening/re-reading a PDF for every template is wasteful.
        self._document_cache = OrderedDict()
        self._document_cache_limit = 8
        self._template_cache = {}
        self._ocr_template_cache = {}
        self._ocr_errors_reported = set()

        # State used by the human-confirmed OCR workflow.
        self.ocr_mode = False
        self.ocr_template_ready = False
        self.ocr_match = None

    def refresh_config(self):
        with open(os.path.join(BASE_DIR, "config.json"), "r", encoding="utf-8") as f:
            self.config = json.load(f)
        template_parent = os.path.dirname(str(self.config.get("TEMPLATE_FOLDER", "") or "")) or BASE_DIR
        test_template_parent = os.path.dirname(str(self.config.get("TEST_TEMPLATE_FOLDER", "") or "")) or BASE_DIR
        self.config.setdefault("OCR_TEMPLATE_FOLDER", os.path.join(template_parent, "OCR_Templates"))
        self.config.setdefault("TEST_OCR_TEMPLATE_FOLDER", os.path.join(test_template_parent, "OCR_Templates"))
        self.config.setdefault("OCR_FUZZY_THRESHOLD", 0.72)
        self.config.setdefault("OCR_DPI", 250)
        self.config.setdefault("OCR_LANGUAGE", "eng")
        self.config.setdefault("TESSDATA_PREFIX", "")
        self.config.setdefault("PRINTER_NAME", "")
        self.config.setdefault("MIN_EMBEDDED_TEXT_CHARS", 40)
        self._template_cache.clear()
        self._ocr_template_cache.clear()

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
        else:
            folder = os.path.abspath(folder)
            self._template_cache.pop(folder, None)
            self._ocr_template_cache.pop(folder, None)

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
        invoice_num = str(invoice_num or "").strip()
        if prefix and not invoice_num.startswith(prefix):
            invoice_num = f"{prefix}{invoice_num}"
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
                    ocr_mode=False, ocr_template_folder=None):
        # Away mode intentionally bypasses review, matching the existing behavior.
        if root.AWAY_MODE and not testing:
            self.print_invoice(filepath)
            return [filepath, False]

        self.ocr_mode = bool(ocr_mode)
        self.ocr_match = None
        self.ocr_template_ready = False
        ocr_template_folder = ocr_template_folder or self.config.get("OCR_TEMPLATE_FOLDER")

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
                    self.log(f"OCR template required for {filename}", tag="yellow", display=True)
                    initial_text = ""
                    if not testing:
                        self.send_email("Must create OCR template", root)

                rectangulator, text_box = self.open_rectangulator(
                    filepath, ocr_template_folder, root,
                    scanner=(testing == "scanner"),
                    ocr_mode=True,
                    initial_text=initial_text,
                    ocr_match=self.ocr_match,
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

    def open_rectangulator(self, pdf_path, template_folder, root, scanner=False, ocr_mode=False, initial_text="", ocr_match=None):  # Setup the page for the Rectangulator and return the Rectangulator and textbox
        # Reset flags
        self.should_print = True 
        self.should_save = True
        self.hit_submit = False
        self.invoice = True
        self.current_page = 0
        self.ocr_mode = bool(ocr_mode)
        self.ocr_match = ocr_match

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


        # Create a Not An Invoice button
        not_inv_button_ax = self.fig.add_axes([0.64, 0.005, 0.18, 0.075])
        not_inv_button = Button(not_inv_button_ax, "Not An Invoice")

        def not_invoice(event):  # If the user clicks the "Not Invoice" button
            try:
                self.invoice = False
                self.ax.clear()
                self.ax.axis("off")
                self.fig.canvas.draw_idle()
                self.done_var.set(1)
            except Exception as e:
                self.log(f"Error in not_invoice: {str(e)} \n{traceback.format_exc()}")
                self.done_var.set(1) 

        not_inv_button.on_clicked(not_invoice)

        # Create a checkbox for if it should be printed
        print_checkbox_ax = self.fig.add_axes([0.86, 0.03, 0.03, 0.03])
        print_checkbox = CheckButtons(print_checkbox_ax, [""], [self.should_print])
        def print_callback(label):
            self.should_print = not self.should_print
        print_checkbox.on_clicked(print_callback)
        # Set the checkbox to be a square and centered
        for i, line in enumerate(print_checkbox.lines):
            rect = print_checkbox.rectangles[i]
            rect.set_width(1)
            rect.set_height(1)
            rect.set_edgecolor("none")
            # Calculate the center of the rectangle
            center_x = rect.get_width() / 2
            center_y = rect.get_height() / 2
            # Update the line positions to be centered
            line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
            line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])
        print_label = self.fig.text(0.853, 0.075, "Print?", fontsize=9)

        # Create a checkbox for if it should be saved
        save_checkbox_ax = self.fig.add_axes([0.93, 0.03, 0.03, 0.03])
        save_checkbox = CheckButtons(save_checkbox_ax, [""], [True])
        def save_callback(label):
            self.should_save = not self.should_save
        save_checkbox.on_clicked(save_callback)
        # Set the checkbox to be a square and centered
        for i, line in enumerate(save_checkbox.lines):
            rect = save_checkbox.rectangles[i]
            rect.set_width(1)
            rect.set_height(1)
            rect.set_edgecolor("none")
            # Calculate the center of the rectangle
            center_x = rect.get_width() / 2
            center_y = rect.get_height() / 2
            # Update the line positions to be centered
            line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
            line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])
        save_label = self.fig.text(0.92, 0.075, "Save?", fontsize=9)

        # Filename text box and submit button
        text_box_ax = self.fig.add_axes([0.1, 0.005, 0.35, 0.075])
        text_box = TextBox(text_box_ax, label="", initial=initial_text or "")
        try:
            text_box.set_active(True)
        except Exception as e:
            self.log(f"Error activating text box: {str(e)} \n{traceback.format_exc()}")
        text_box.text_disp.set_horizontalalignment('right')
        text_box.text_disp.set_position((0.95, 0.5))

        # 3. Clip the text so anything spilling out the left side is visually hidden
        text_box.text_disp.set_clip_on(True)
        text_box.text_disp.set_clip_box(text_box_ax.bbox)

        def on_text_submit(event=None):
            # Runs on the Tk main thread (matplotlib button callback). create_alert
            # uses wait_window(), which MUST run on the main thread — do not spawn
            # a worker thread here.
            if self.hit_submit:
                return
            self.hit_submit = True
            try:
                entered = self.sanitize_filename(text_box.text)
                if not entered:
                    self.create_alert("Please enter a filename before submitting.")
                    return
                if self.ocr_mode and not self.ocr_template_ready:
                    self.create_alert("Please draw and verify the three OCR rectangles before submitting.")
                    return
                if entered != text_box.text:
                    text_box.set_val(entered)
                filename_is_correct = self.create_alert(
                    f"Is '{entered}' the correct filename?")
                if filename_is_correct:
                    self.ax.clear()
                    self.ax.axis("off")
                    self.fig.canvas.draw_idle()
                    self.done_var.set(1)  # signal that the user is done
            finally:
                self.hit_submit = False

        submit_button_ax = self.fig.add_axes([0.45, 0.005, 0.15, 0.075])
        submit_button = Button(submit_button_ax, "Submit")
        submit_button.on_clicked(on_text_submit)

        # Create text labels for instructions and text box
        text_box_label = self.fig.text(
            0.1,
            0.089,
            "Enter Filename Manually (mm-dd-yy_invoice#)",
            fontsize=10)
        if self.ocr_mode and ocr_match:
            confidence = ocr_match.get("score", 0.0) * 100
            line1 = f"- OCR template match: {confidence:.1f}% confidence — review the suggested filename below"
            line2 = "- OCR never renames automatically; edit the suggestion if needed, then click Submit"
            line3 = "- You may draw three new rectangles and right-click if this OCR template needs retraining"
        elif self.ocr_mode:
            line1 = "- OCR mode: draw boxes around Company Name, Date, and Invoice (in that order)"
            line2 = "- Right-click to OCR/verify the selections and create a separate OCR template"
            line3 = "- The OCR attempt will fill the filename box; review/edit it, then click Submit"
        else:
            line1 = "- Left Click to Draw boxes around Company Name, Date, and Invoice (in that order)"
            line2 = "- Company Name can be any piece of text unique to that vendor"
            line3 = "- Right Click to verify and save, Middle Click to Pan, Scroll to Zoom"
        instruction_label = self.fig.text(0.1, 0.975, line1, fontsize=10)
        instruction_label_2 = self.fig.text(0.1, 0.95, line2, fontsize=10)
        instruction_label_3 = self.fig.text(0.1, 0.925, line3, fontsize=10)
        
        # Create next and previous page buttons
        prev_button_ax = self.fig.add_axes([0.02, 0.45, 0.04, 0.1])
        prev_button = Button(prev_button_ax, "<")
        def on_prev(event):
            if self.current_page > 0:
                self.current_page -= 1
                draw_page(self.current_page)
        prev_button.on_clicked(on_prev)

        next_button_ax = self.fig.add_axes([0.94, 0.45, 0.04, 0.1])
        next_button = Button(next_button_ax, ">")
        def on_next(event):
            if self.current_page < self.total_pages - 1:
                self.current_page += 1
                draw_page(self.current_page)
        next_button.on_clicked(on_next)

        self.fig.canvas.draw_idle()
        rectangulator = Rectangulator(self.ax, self.fig, pdf_path, template_folder, self, text_box=text_box, ocr_mode=self.ocr_mode)
        self.rectangulator_instance = rectangulator
        rectangulator.page_num = self.current_page
        root.root.wait_variable(self.done_var)  # wait for the user to finish

        # Remove the text labels and rectangles
        for ax in [text_box_ax, not_inv_button_ax, print_checkbox_ax, save_checkbox_ax, submit_button_ax, prev_button_ax, next_button_ax]:
            self.fig.delaxes(ax)
        for label in [text_box_label, instruction_label, instruction_label_2, instruction_label_3, print_label, save_label]:
            label.remove()
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

    def get_vendor_prefix(self, identifier):  # Get the prefix for a vendor if it exists in the config
        # Parses pairs like "VendorA:VA-, VendorB:VB-"
        prefix_str = self.config.get("PREFIX_VENDORS", "")
        if not prefix_str:
            return ""
            
        pairs = [p.strip() for p in prefix_str.split(",") if ":" in p]
        for pair in pairs:
            vendor, prefix = pair.split(":", 1)
            if vendor.strip() == identifier.strip():
                return prefix.strip()
        return ""

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
        date_patterns = [
            "%B %d, %Y",
            "%b %d, %Y",
            "%B %d, %y",
            "%b %d, %y",
            "%d-%B-%Y",
            "%d-%b-%Y",
            "%d-%B-%y",
            "%d-%b-%y",
            "%m-%d-%Y",
            "%m-%d-%y",
            "%b %d %Y",
            "%B %d %Y",
            "%b %d %y",
            "%B %d %y",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%Y-%m-%d",
            "%y-%m-%d",
            "%Y/%m/%d",
            "%y/%m/%d"
        ]
        invoice_date = str(invoice_date).replace("/", "-")
        for pattern in date_patterns:
            try:
                dt = datetime.strptime(invoice_date, pattern)
                return dt.strftime("%m-%d-%y")
            except ValueError:
                continue
        self.log(f"Could not convert {invoice_date} to date")
        return invoice_date

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

    def __init__(self, ax, fig, pdf_path, template_folder, rectangulator_handler, text_box=None, ocr_mode=False):
        self.rectangulator_handler = rectangulator_handler
        self.pdf_path = pdf_path
        self.template_folder = template_folder
        self.fig = fig
        self.ax = ax
        self.text_box = text_box
        self.ocr_mode = bool(ocr_mode)

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

    def on_key_press(self, event):  # Handle key press events
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

            # Start drawing a rectangle
            self.start_x = event.xdata
            self.start_y = event.ydata
            self.rect = Rectangle((self.start_x, self.start_y), 0, 0, edgecolor="red", linewidth=2, fill=False)
            self.ax.add_patch(self.rect)
            self.ax.figure.canvas.draw()

        elif event.button == 2:  # middle mouse button, pan
            self.pan_start = (event.x, event.y)

        elif event.button == 3:  # right mouse button, save rectangles
            # Check if 3 rectangles have been drawn
            if len(self.rectangles) != 3:
                self.rectangulator_handler.log(
                    "Please draw exactly three rectangles", display=True)
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

        headers = ["--- Company Name: ", "--- Invoice Date: ", "--- Invoice Number: "]
        extracted_values = []
        for rect in self.rectangles:
            extracted_values.append(
                self.rectangulator_handler.get_text_in_rect(
                    rect, self.pdf_path, self.page_num, ocr=self.ocr_mode
                )
            )
        extracted_text = "\n".join(h + v for h, v in zip(headers, extracted_values))
        text_is_correct = self.rectangulator_handler.create_alert(
            f"Does the following text match what you selected?\n\n{extracted_text}",
            numbered_buttons=3,
        )

        if isinstance(text_is_correct, int) and not isinstance(text_is_correct, bool):
            self.correcting_rect_index = text_is_correct
            self.rectangulator_handler.log(f"Please reselect {headers[self.correcting_rect_index]}")
            self.reset_rectangles(specific_rect=self.correcting_rect_index)
            return

        if text_is_correct:
            if self.ocr_mode:
                if not all(extracted_values):
                    self.rectangulator_handler.log("OCR did not read all three selections. Please redraw the empty field.", tag="orange")
                    return
                if self.save_template(ocr=True):
                    suggestion = self.rectangulator_handler.build_filename_from_fields(*extracted_values)
                    self.rectangulator_handler.ocr_template_ready = True
                    if self.text_box is not None:
                        self.text_box.set_val(suggestion)
                    self.rectangulator_handler.log(
                        f"OCR suggested '{suggestion}'. Review/edit the filename and click Submit.",
                        tag="yellow", display=True,
                    )
                    # Keep the window open: the user is always the final authority.
                    self.reset_rectangles()
                return

            self.ax.clear()
            self.ax.axis("off")
            self.fig.canvas.draw_idle()
            self.rectangulator_handler.done_var.set(1)
            return

        self.rectangulator_handler.log("Please reselect rectangles", display=True)
        self.reset_rectangles()

