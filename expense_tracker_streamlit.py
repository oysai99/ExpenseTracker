"""
Streamlit Expense Tracker — integrated Amex-specific parser + OCR/generic extractor
"""

import streamlit as st
import tempfile
import os
import fitz  # PyMuPDF
import re
import pdfplumber
import pandas as pd
from datetime import datetime
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from typing import List, Dict, Optional

# Optional OCR imports (import lazily and handle missing packages)
OCR_AVAILABLE = True
try:
    import pytesseract
    from pdf2image import convert_from_path
except Exception:
    OCR_AVAILABLE = False

# ------------------------ Classification ------------------------
DEFAULT_MAPPING = {
    'food': ['mc', 'starbucks', 'coffee', 'restaurant', 'dining', 'kfc', 'burger', 'mcdonald', 'pizza', 'eat', 'cafe'],
    'transport': ['grab', 'uber', 'taxi', 'ezlink', 'bus', 'train', 'metro', 'transport'],
    'hotels': ['hotel', 'booking', 'airbnb', 'marriott', 'hilton', 'agoda'],
    'air tickets': ['airlines', 'flight', 'sq', 'airasia', 'emirates', 'cathay', 'ticket'],
    'shopping': ['amazon', 'shopee', 'lazada', 'uniqlo', 'zara', 'apple', 'hm', 'store'],
    'utilities': ['singtel', 'starhub', 'netflix', 'spotify', 'electricity', 'water', 'gas', 'utility'],
    'health': ['clinic', 'hospital', 'doctor', 'pharmacy'],
    'salary': ['salary', 'payroll'],
}

def classify_description(desc: str, mapping: Dict[str, List[str]] = DEFAULT_MAPPING) -> str:
    desc = (desc or "").lower()
    for cat, keys in mapping.items():
        for k in keys:
            if k in desc:
                return cat
    return 'misc'

# ------------------------ Helpers ------------------------
def standardize_amount(s: str) -> Optional[float]:
    if not s:
        return None
    s = s.strip()
    s = s.replace(',', '')
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1]
    s = re.sub(r'[^\d\.\-]', '', s)
    try:
        val = float(s)
        return -val if negative else val
    except Exception:
        return None

def detect_iso_currency_from_text(text: str) -> Optional[str]:
    codes = re.findall(r'\b([A-Z]{3})\b', text)
    if codes:
        codes = [c.upper() for c in codes]
        return max(set(codes), key=codes.count)
    return None

def detect_currency_symbol(text: str) -> Optional[str]:
    if 's$' in text.lower() or 'sgd' in text.lower():
        return 'SGD'
    if '$' in text and 'aud' not in text.lower() and 'cad' not in text.lower():
        return 'USD'
    if '£' in text or 'gbp' in text.lower():
        return 'GBP'
    if '€' in text or 'eur' in text.lower():
        return 'EUR'
    return None

# ------------------------ PDF Screenshot Helper ------------------------
def crop_bbox_from_pdf(pdf_path: str, page_no: int, bbox, outpath: str, zoom: float = 2.0) -> bool:
    """
    Crop given bbox (x0, top, x1, bottom) in PDF page coordinates using PyMuPDF.
    """
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_no]
        x0, y0, x1, y1 = bbox
        pad = 2
        rect = fitz.Rect(max(0, x0 - pad), max(0, y0 - pad), x1 + pad, y1 + pad)
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(clip=rect, matrix=matrix, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(outpath, format='PNG')
        return True
    except Exception:
        return False

def crop_bbox_from_image(image: Image.Image, bbox, outpath: str) -> bool:
    """
    Crop bbox on a PIL Image. bbox is (x0, y0, x1, y1).
    """
    try:
        left, top, right, bottom = bbox
        cropped = image.crop((left, top, right, bottom))
        cropped.save(outpath, format='PNG')
        return True
    except Exception:
        return False

# ------------------------ Text grouping ------------------------
def group_words_to_lines(words: List[Dict]) -> List[Dict]:
    lines = {}
    for w in words:
        top = round(w['top'], 1)
        key = None
        for k in lines.keys():
            if abs(k - top) <= 2:
                key = k
                break
        if key is None:
            key = top
            lines[key] = {'text': w['text'], 'x0': w['x0'], 'x1': w['x1'], 'top': w['top'], 'bottom': w['bottom'], 'words': [w]}
        else:
            lines[key]['text'] += ' ' + w['text']
            lines[key]['x0'] = min(lines[key]['x0'], w['x0'])
            lines[key]['x1'] = max(lines[key]['x1'], w['x1'])
            lines[key]['bottom'] = max(lines[key]['bottom'], w['bottom'])
            lines[key]['words'].append(w)
    return [lines[k] for k in sorted(lines.keys())]

# ------------------------ OCR helper (raster-based) ------------------------
def ocr_page_to_words(page_image: Image.Image) -> List[Dict]:
    """
    Use pytesseract.image_to_data to extract words with positions from a PIL Image.
    Returns list of dicts with keys: text, x0, top, x1, bottom.
    """
    if not OCR_AVAILABLE:
        return []
    data = pytesseract.image_to_data(page_image, output_type=pytesseract.Output.DICT)
    words = []
    n = len(data['level'])
    for i in range(n):
        text = data['text'][i].strip()
        if not text:
            continue
        x = int(data['left'][i])
        y = int(data['top'][i])
        w = int(data['width'][i])
        h = int(data['height'][i])
        words.append({'text': text, 'x0': x, 'x1': x + w, 'top': y, 'bottom': y + h})
    return words

# ------------------------ Transaction Extractor (generic) ------------------------
DATE_PATTERNS = [
    r'\b(\d{4}-\d{2}-\d{2})\b',
    r'\b(\d{2}/\d{2}/\d{4})\b',
    r'\b(\d{2}-\d{2}-\d{4})\b',
    r'\b(\d{1,2}\s+[A-Za-z]{3,}\s+\d{4})\b'
]

AMOUNT_REGEX = re.compile(r'([A-Z]{3})?\s*([\(\d][\d,]*\.\d{2}|\([\d,]*\.\d{2}\)')

def extract_transactions_from_pdf(pdf_path: str, bank_hint: str = 'Auto', use_ocr_if_empty: bool = False) -> List[Dict]:
    """
    Generic extractor. Uses pdfplumber first; if use_ocr_if_empty and a page has no words,
    and OCR is available, it rasterizes the page and extracts words using pytesseract.
    """
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_no, page in enumerate(pdf.pages):
                words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
                used_ocr = False
                image_for_ocr = None
                if (not words or len(words) == 0) and use_ocr_if_empty and OCR_AVAILABLE:
                    # rasterize page to image and OCR
                    try:
                        pil_pages = convert_from_path(pdf_path, first_page=page_no+1, last_page=page_no+1, dpi=200)
                        if pil_pages:
                            page_image = pil_pages[0]
                            image_for_ocr = page_image
                            ocr_words = ocr_page_to_words(page_image)
                            # convert image pixel coords into pseudo-PDF coords by keeping them as-is but note cropping uses image
                            words = ocr_words
                            used_ocr = True
                    except Exception:
                        used_ocr = False

                if not words:
                    continue
                lines = group_words_to_lines(words)
                for line in lines:
                    text = line['text'].strip()
                    if not re.search(r'\d+\.\d{2}', text):
                        continue
                    date = None
                    for dp in DATE_PATTERNS:
                        m = re.search(dp, text)
                        if m:
                            date = m.group(1)
                            break
                    amounts = AMOUNT_REGEX.findall(text)
                    token_amounts = re.findall(r'[\(\d][\d,]*\.\d{2}|\([\d,]*\.\d{2}\)', text)
                    amount_card = None
                    currency_orig = None
                    amount_orig = None
                    card_currency = None

                    if amounts and len(amounts) >= 1:
                        if len(amounts) == 1:
                            iso, amt = amounts[0]
                            amount_card = standardize_amount(amt)
                            if iso:
                                card_currency = iso
                        else:
                            iso1, amt1 = amounts[0]
                            iso2, amt2 = amounts[1]
                            amount_orig = standardize_amount(amt1)
                            amount_card = standardize_amount(amt2)
                            currency_orig = iso1 if iso1 else currency_orig
                            card_currency = iso2 if iso2 else card_currency

                    if not amount_card and token_amounts:
                        amount_card = standardize_amount(token_amounts[-1])

                    if not card_currency:
                        card_currency = detect_iso_currency_from_text(text) or detect_currency_symbol(text)
                    if not currency_orig:
                        iso_codes = re.findall(r'\b([A-Z]{3})\b', text)
                        if iso_codes:
                            currency_orig = iso_codes[0] if len(iso_codes) == 1 else iso_codes[0]

                    description = text
                    category = classify_description(description)

                    # compute bbox for cropping
                    bbox = (line['x0'], line['top'], line['x1'], line['bottom'])
                    txn = {
                        'date': date,
                        'description': description,
                        'currency_orig': currency_orig,
                        'amount_orig': amount_orig,
                        'card_currency': card_currency,
                        'amount_card': amount_card,
                        'category': category,
                        'page': page_no,
                        'bbox': bbox,
                        'used_ocr': used_ocr,
                        'page_image_for_ocr': image_for_ocr  # only set in-memory; not serializable
                    }
                    transactions.append(txn)
    except Exception:
        # return empty; caller will handle fallback
        return []
    return transactions

# ------------------------ Amex-specific extractor ------------------------
_AMEX_CURRENCY_NAME_TO_ISO = {
    'UNITED STATES DOLLAR': 'USD',
    'UNITED STATES DOLLARS': 'USD',
    'UNITED STATES': 'USD',
    'THAILAND BAHT': 'THB',
    'THAI BAHT': 'THB',
    'PHILIPPINE PESO': 'PHP',
    'PHILIPPINE PESOS': 'PHP',
    'CHINA YUAN RENMINBI': 'CNY',
    'CHINESE YUAN': 'CNY',
    'SINGAPORE DOLLAR': 'SGD',
    'SGD': 'SGD',
    'USD': 'USD',
    'EUR': 'EUR',
    'POUND STERLING': 'GBP',
    'GBP': 'GBP',
}

_DATE_DOT_PATTERN = re.compile(r'^\s*(\d{2}\.\d{2}\.\d{2})\b')  # e.g. 23.06.24

def _map_currency_name_to_iso(name: str) -> Optional[str]:
    if not name:
        return None
    s = re.sub(r'[^A-Za-z\s]', ' ', name).upper().strip()
    if s in _AMEX_CURRENCY_NAME_TO_ISO:
        return _AMEX_CURRENCY_NAME_TO_ISO[s]
    for k, v in _AMEX_CURRENCY_NAME_TO_ISO.items():
        if k in s:
            return v
    iso = detect_iso_currency_from_text(s)
    if iso:
        return iso
    return None

def parse_amex_line(line: Dict, next_line: Optional[Dict], page_width: float) -> Optional[Dict]:
    """
    Parse a grouped line dict (as returned by group_words_to_lines) for Amex layout.
    Returns a transaction dict or None if insufficient data.
    This helper is unit-testable with synthetic line/word dicts.
    """
    text = line.get('text', '').strip()
    mdate = _DATE_DOT_PATTERN.match(text)
    if not mdate:
        return None
    date_token = mdate.group(1)
    line_words = line.get('words', [])
    if not line_words:
        return None

    right_col_x = page_width * 0.78
    foreign_col_x = page_width * 0.60

    description_parts = []
    for w in line_words:
        if re.match(r'^\d{2}\.\d{2}\.\d{2}$', w.get('text', '').strip()):
            continue
        if w.get('x0', 0) >= right_col_x:
            continue
        description_parts.append(w.get('text', ''))
    description = ' '.join(description_parts).strip()

    num_pat = re.compile(r'[\(\d][\d,]*\.\d{2}|\([\d,]*\.\d{2}\)')
    right_tokens = [w for w in line_words if w.get('x0', 0) >= right_col_x]
    amount_card = None
    for w in reversed(right_tokens):
        if num_pat.search(w.get('text', '')):
            amount_card = standardize_amount(w.get('text'))
            break

    amount_orig = None
    currency_orig = None
    foreign_tokens_same = [w for w in line_words if foreign_col_x <= w.get('x0', 0) < right_col_x]
    if foreign_tokens_same:
        num = None
        name_parts = []
        for w in foreign_tokens_same:
            t = w.get('text', '')
            if num_pat.search(t) and num is None:
                num = t
            else:
                name_parts.append(t)
        if num:
            amount_orig = standardize_amount(num)
        if name_parts:
            currency_orig = _map_currency_name_to_iso(' '.join(name_parts))

    if (amount_orig is None or currency_orig is None) and next_line:
        next_words = next_line.get('words', [])
        next_foreign_tokens = [w for w in next_words if w.get('x0', 0) >= foreign_col_x]
        if next_foreign_tokens:
            num = None
            name_parts = []
            for w in next_foreign_tokens:
                t = w.get('text', '')
                if num_pat.search(t) and num is None:
                    num = t
                else:
                    name_parts.append(t)
            if num and amount_orig is None:
                amount_orig = standardize_amount(num)
            if name_parts and currency_orig is None:
                currency_orig = _map_currency_name_to_iso(' '.join(name_parts))

    txn = {
        'date': date_token,
        'description': description,
        'currency_orig': currency_orig,
        'amount_orig': amount_orig,
        'card_currency': None,
        'amount_card': amount_card,
        'category': classify_description(description),
        'page': None,
        'bbox': (line.get('x0'), line.get('top'), line.get('x1'), line.get('bottom'))
    }

    if txn['amount_card'] is None and txn['amount_orig'] is None:
        return None
    return txn

def extract_amex_specific(pdf_path: str) -> List[Dict]:
    txns = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_no, page in enumerate(pdf.pages):
                words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
                if not words:
                    continue
                page_width = page.width
                lines = group_words_to_lines(words)
                for i, line in enumerate(lines):
                    next_line = lines[i+1] if i+1 < len(lines) else None
                    parsed = parse_amex_line(line, next_line, page_width)
                    if parsed:
                        parsed['page'] = page_no
                        txns.append(parsed)
    except Exception:
        return []
    return txns

# fallback wrapper kept for compatibility
def extract_amex(pdf_path, use_ocr_if_empty=False):
    txns = extract_amex_specific(pdf_path)
    return txns or [{
        "date": "2025-10-28",
        "description": "Starbucks Coffee",
        "currency_orig": "USD",
        "amount_orig": 5.50,
        "card_currency": "SGD",
        "amount_card": 7.50,
        "page": 0,
        "bbox": (0,0,500,50),
        "used_ocr": False
    }]

# Simple wrappers for other banks (keep generic extractor)
def extract_citibank(pdf_path, use_ocr_if_empty=False):
    return extract_transactions_from_pdf(pdf_path, bank_hint='Citibank', use_ocr_if_empty=use_ocr_if_empty)

def extract_posb(pdf_path, use_ocr_if_empty=False):
    return extract_transactions_from_pdf(pdf_path, bank_hint='POSB', use_ocr_if_empty=use_ocr_if_empty)

# ------------------------ Auto currency detection ------------------------
def detect_base_currency(pdf_path: str) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ''
            for i in range(min(2, len(pdf.pages))):
                text += '\n' + (pdf.pages[i].extract_text() or '')
            iso = detect_iso_currency_from_text(text)
            if iso:
                return iso
            sym = detect_currency_symbol(text)
            if sym:
                return sym
    except Exception:
        pass
    return 'SGD'

# ------------------------ Excel Writer ------------------------
def create_excel(transactions, outpath):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Transactions'
    headers = ['Date','Description','Currency (orig)','Amount (orig)','Card Currency',
               'Amount (card)','Category','Base Currency','Screenshot']
    ws.append(headers)

    for i, t in enumerate(transactions, start=2):
        ws.append([
            t.get('date'), t.get('description'), t.get('currency_orig'), t.get('amount_orig'),
            t.get('card_currency'), t.get('amount_card'), t.get('category'),
            t.get('base_currency'), ''
        ])
        if t.get('img_path') and os.path.exists(t['img_path']):
            try:
                img = XLImage(t['img_path'])
                ws.add_image(img, f'I{i}')
            except Exception:
                pass
    wb.save(outpath)
    return outpath

# ------------------------ Streamlit App ------------------------
st.set_page_config(page_title='Expense Tracker', layout='wide')
st.title('🏦 Expense Tracker — PDF Statement to Excel (Amex parser)')

bank_type = st.sidebar.selectbox('Select Bank Format', ['Auto-detect', 'Amex', 'Citibank', 'POSB'])
uploaded_files = st.file_uploader('Upload statement PDFs', type=['pdf'], accept_multiple_files=True)
preview_limit = st.sidebar.number_input('Preview max transactions per file', min_value=5, max_value=200, value=50)
enable_ocr = st.sidebar.checkbox('Enable OCR for scanned PDFs (requires tesseract & poppler)', value=False)

if enable_ocr and not OCR_AVAILABLE:
    st.warning('OCR requested but pytesseract/pdf2image are not installed in the environment. Please install them and ensure system packages tesseract-ocr and poppler are available.')

if uploaded_files and st.button('Process and Generate Excel'):
    tmpdir = tempfile.mkdtemp(prefix='expense_tracker_')
    all_txns = []

    for up in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', dir=tmpdir) as tf:
            tf.write(up.read())
            pdf_path = tf.name

        base_curr = detect_base_currency(pdf_path)
        st.write(f'📄 **{up.name}** — Detected Base Currency: **{base_curr}**')

        # Choose extractor with OCR fallback if enabled
        if bank_type == 'Amex':
            txns = extract_amex_specific(pdf_path)
        elif bank_type == 'Citibank':
            txns = extract_citibank(pdf_path, use_ocr_if_empty=enable_ocr)
        elif bank_type == 'POSB':
            txns = extract_posb(pdf_path, use_ocr_if_empty=enable_ocr)
        else:
            txns = extract_transactions_from_pdf(pdf_path, use_ocr_if_empty=enable_ocr)

        if not txns:
            st.warning(f'No transactions found in {up.name}. Using fallback sample.')
            txns = extract_amex(pdf_path, use_ocr_if_empty=enable_ocr)

        for idx, t in enumerate(txns):
            if idx >= preview_limit:
                break
            t['category'] = classify_description(t.get('description'))
            t['base_currency'] = base_curr
            # For Amex, set card_currency to detected base currency
            if bank_type == 'Amex':
                t['card_currency'] = base_curr
            page_no = t.get('page', 0) or 0
            bbox = t.get('bbox', None)
            img_path = None
            if bbox:
                safe_name = re.sub(r'[^\\w\-_\. ]', '_', up.name)
                img_path = os.path.join(tmpdir, f"{safe_name}_p{page_no}_{idx}.png")
                if t.get('used_ocr') and t.get('page_image_for_ocr') is not None:
                    ok = crop_bbox_from_image(t['page_image_for_ocr'], bbox, img_path)
                    if not ok:
                        img_path = None
                else:
                    ok = crop_bbox_from_pdf(pdf_path, page_no, bbox, img_path)
                    if not ok:
                        img_path = None
            t['img_path'] = img_path
            if 'page_image_for_ocr' in t:
                t.pop('page_image_for_ocr', None)
            all_txns.append(t)

    if not all_txns:
        st.error('No transactions extracted from any uploaded PDFs.')
    else:
        df = pd.DataFrame(all_txns)
        st.dataframe(df[['date','description','currency_orig','amount_orig','card_currency','amount_card','category','base_currency']].head(500))
        out_xlsx = os.path.join(tmpdir, f'expenses_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        create_excel(all_txns, out_xlsx)
        with open(out_xlsx, 'rb') as f:
            st.download_button('Download Excel with Screenshots', data=f.read(), file_name=os.path.basename(out_xlsx))
else:
    st.info('Upload PDF statements and choose your bank format to begin. Enable OCR for scanned PDFs (requires tesseract & poppler).')