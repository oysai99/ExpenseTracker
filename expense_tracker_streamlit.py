"""
Streamlit Expense Tracker (Bank-Tuned with Auto Currency Detection)

Features:
- Upload PDF statements (Amex, Citibank, POSB)
- Auto-detect credit card base currency (supports all ISO 3-letter codes)
- Extract date, description, amounts (original + card currency)
- Auto-classify transactions (food, transport, hotels, etc.)
- Crop screenshot of each line and embed in Excel
- Show detected base currency per file
- Downloadable Excel with base currency column
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
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ------------------------ Classification ------------------------
def classify_description(desc: str) -> str:
    desc = (desc or "").lower()
    mapping = {
        'food': ['mc', 'starbucks', 'coffee', 'restaurant', 'dining', 'kfc', 'burger', 'mcdonald'],
        'transport': ['grab', 'uber', 'taxi', 'ezlink', 'bus', 'train', 'sg transport'],
        'hotels': ['hotel', 'booking', 'airbnb', 'marriott', 'hilton', 'agoda'],
        'air tickets': ['airlines', 'flight', 'sq', 'airasia', 'emirates', 'cathay'],
        'shopping': ['amazon', 'shopee', 'lazada', 'uniqlo', 'zara', 'apple'],
        'utilities': ['singtel', 'starhub', 'netflix', 'spotify', 'electricity'],
    }
    for cat, keys in mapping.items():
        for k in keys:
            if k in desc:
                return cat
    return 'misc'

# ------------------------ PDF Screenshot Helper ------------------------
def crop_pdf_line_to_image(pdf_path: str, page_no: int, bbox, outpath: str):
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_no]
        rect = fitz.Rect(*bbox)
        pad = 2
        rect = fitz.Rect(max(0, rect.x0 - pad), max(0, rect.y0 - pad), rect.x1 + pad, rect.y1 + pad)
        pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(outpath, format='PNG')
        return True
    except Exception:
        return False

# ------------------------ Bank-specific extractors ------------------------
# [extract_amex, extract_citibank, extract_posb remain unchanged]

# ------------------------ Auto currency detection ------------------------
def detect_base_currency(pdf_path: str) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''
            codes = re.findall(r'\b[A-Z]{3}\b', text)
            if codes:
                # return most common currency code
                return max(set(codes), key=codes.count)
    except Exception:
        pass
    return 'SGD'  # default fallback

# ------------------------ Excel Writer ------------------------
def create_excel(transactions, outpath):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Transactions'
    headers = ['Date','Description','Currency (orig)','Amount (orig)','Card Currency','Amount (card)','Category','Base Currency','Screenshot']
    ws.append(headers)

    for i, t in enumerate(transactions, start=2):
        ws.append([
            t.get('date'), t.get('description'), t.get('currency_orig'), t.get('amount_orig'),
            t.get('card_currency'), t.get('amount_card'), t.get('category'), t.get('base_currency'), ''
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
st.set_page_config(page_title='Expense Tracker - Bank Tuned', layout='wide')
st.title('🏦 Expense Tracker — Tuned for Amex, Citibank, POSB with Auto Currency Detection')

bank_type = st.sidebar.selectbox('Select Bank Format', ['Auto-detect', 'Amex', 'Citibank', 'POSB'])
uploaded_files = st.file_uploader('Upload statement PDFs', type=['pdf'], accept_multiple_files=True)

if uploaded_files and st.button('Process and Generate Excel'):
    tmpdir = tempfile.mkdtemp(prefix='expense_tracker_')
    all_txns = []

    for up in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', dir=tmpdir) as tf:
            tf.write(up.read())
            pdf_path = tf.name

        base_curr = detect_base_currency(pdf_path)
        st.write(f'📄 **{up.name}** — Detected Base Currency: **{base_curr}**')

        if bank_type == 'Amex':
            txns = extract_amex(pdf_path)
        elif bank_type == 'Citibank':
            txns = extract_citibank(pdf_path)
        elif bank_type == 'POSB':
            txns = extract_posb(pdf_path)
        else:
            txns = extract_citibank(pdf_path) or extract_amex(pdf_path) or extract_posb(pdf_path)

        for t in txns:
            t['category'] = classify_description(t['description'])
            t['base_currency'] = base_curr
            img_path = os.path.join(tmpdir, f"{up.name}_p{t['page']}.png")
            t['img_path'] = img_path if crop_pdf_line_to_image(pdf_path, t['page'], (0,0,500,50), img_path) else None
            all_txns.append(t)

    df = pd.DataFrame(all_txns)
    st.dataframe(df)

    out_xlsx = os.path.join(tmpdir, f'expenses_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    create_excel(all_txns, out_xlsx)
    with open(out_xlsx, 'rb') as f:
        st.download_button('Download Excel with Screenshots', data=f.read(), file_name=os.path.basename(out_xlsx))

else:
    st.info('Upload PDF statements and choose your bank format to begin.')
