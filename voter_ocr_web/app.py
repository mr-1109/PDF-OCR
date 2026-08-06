import os
import io
import base64
import json
import re
import shutil
import time
import uuid
from flask import Flask, render_template, request, jsonify, send_file, Response
from werkzeug.utils import secure_filename
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import fitz  # PyMuPDF

app = Flask(__name__)
app.secret_key = os.urandom(16)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AMIT_RJ_DIR = os.path.abspath(os.path.join(BASE_DIR, '..', '..', '..'))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['EXCEL_OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'excel_outputs')
app.config['AREA_EXCEL_OUTPUT_FOLDER'] = os.path.join(AMIT_RJ_DIR, 'processed excel with area')
app.config['PROCESSED_PDF_FOLDER'] = os.path.join(BASE_DIR, 'processed_pdfs')
app.config['MAX_CONTENT_LENGTH'] = None
BATCH_OUTPUT_DIRS = {}
BATCH_OUTPUT_FOLDER_NAME = 'processed excel'
LOCAL_FOLDER_SEARCH_ROOTS = (
    os.path.expanduser('~/Desktop'),
    os.path.expanduser('~/Downloads'),
)
for folder in (
    app.config['UPLOAD_FOLDER'],
    app.config['EXCEL_OUTPUT_FOLDER'],
    app.config['AREA_EXCEL_OUTPUT_FOLDER'],
    app.config['PROCESSED_PDF_FOLDER'],
):
    os.makedirs(folder, exist_ok=True)

# Hardcoded layout constants
HEADER_HEIGHT = 114
START_GAP = 57
CARD_WIDTH = 792
CARD_HEIGHT = 330
MAX_CELLS = 30

AREA_IGNORE_TERMS = (
    'नाम', 'लिंग', 'मकान', 'आयु', 'Photo', 'Available', 'घटक', 'सूची',
    'पृष्ठ', 'पकष', 'जनवरी', 'जनवरर', 'अनुसार', 'अनपसरर', 'Deleted',
    'नगर', 'निगम', 'परिषद', 'विधान', 'वपरर', 'वार्ड', 'भाग', 'रररर',
)
ACTION_CODES = {'E', 'S', 'R', 'O'}
DEVANAGARI_DIGITS = str.maketrans('०१२३४५६७८९', '0123456789')
AREA_ONLY_EXPORT_COLUMNS = ['Serial', 'List Type', 'Action', 'Area']
FULL_EXPORT_COLUMNS = [
    'WARD_NO', 'PART_NO', 'Page', 'List Type', 'Action', 'Serial', 'EPIC',
    'Name', 'Relation', 'F_NAME', 'House Number', 'Age', 'Gender'
]


def pdf_page_to_image(pdf_path, page_num=0, dpi=200):
    doc = fitz.open(pdf_path)
    try:
        if page_num >= len(doc):
            return None
        page = doc[page_num]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    finally:
        doc.close()


def ocr_image(img):
    raw = pytesseract.image_to_string(img, lang='eng+hin')
    return raw.strip()


def ocr_image_confidence(img):
    data = pytesseract.image_to_data(img, lang='eng+hin', output_type=pytesseract.Output.DICT)
    texts = []
    for i, text in enumerate(data['text']):
        if int(data['conf'][i]) > 20 and text.strip():
            texts.append(text.strip())
    return ' '.join(texts)


def _parse_bool(value):
    return str(value or '').strip().lower() in ('1', 'true', 'yes', 'on')


def _group_close_numbers(values, tolerance=1.2):
    groups = []
    for value in sorted(values):
        if not groups or abs(groups[-1][-1] - value) > tolerance:
            groups.append([value])
        else:
            groups[-1].append(value)
    return [sum(group) / len(group) for group in groups]


def _unique_card_segments(segments):
    unique = []
    for x0, x1 in sorted(segments):
        if not unique or abs(unique[-1][0] - x0) > 8:
            unique.append((x0, x1))
    return unique


def _pdf_block_text(block):
    return ' '.join(
        span.get('text', '')
        for line in block.get('lines', [])
        for span in line.get('spans', [])
    ).strip()


def _count_text_columns(page, top, bottom):
    starts = []
    for block in page.get_text('dict').get('blocks', []):
        if block.get('type') != 0:
            continue
        text = _pdf_block_text(block)
        if 'नरम' not in text and 'नाम' not in text:
            continue
        x0, y0, _x1, _y1 = block['bbox']
        if top + 4 <= y0 <= bottom - 8 and x0 < page.rect.width - 60:
            starts.append(x0)
    return min(len(_group_close_numbers(starts, tolerance=18)), 3)


def _serials_in_row(page, top, bottom):
    serials = []
    for block in page.get_text('dict').get('blocks', []):
        if block.get('type') != 0:
            continue
        text = _pdf_block_text(block)
        cleaned = re.sub(r'\D+', '', text)
        if not cleaned or not re.fullmatch(r'\d{1,5}', cleaned):
            continue
        x0, y0, _x1, _y1 = block['bbox']
        if top <= y0 <= min(bottom, top + 28):
            serials.append((x0, cleaned))
    unique = []
    seen = set()
    for _x, serial in sorted(serials):
        if serial not in seen:
            unique.append(serial)
            seen.add(serial)
    return unique


def _action_from_text(value):
    text = re.sub(r'\s+', ' ', str(value or '').upper()).strip()
    if not text:
        return ''
    matches = re.findall(r'(?<![A-Z0-9])([ESRO])(?![A-Z0-9])', text)
    return matches[-1] if matches else ''


def _card_markers_in_row(page, top, bottom):
    serials = []
    top_blocks = []
    top_band_bottom = min(bottom, top + 36)

    for block in page.get_text('dict').get('blocks', []):
        if block.get('type') != 0:
            continue
        text = _pdf_block_text(block)
        if not text:
            continue
        x0, y0, x1, _y1 = block['bbox']
        if top <= y0 <= top_band_bottom:
            top_blocks.append((x0, x1, text))
            cleaned = re.sub(r'\D+', '', text)
            if cleaned and re.fullmatch(r'\d{1,5}', cleaned):
                serials.append({'x': x0, 'serial': cleaned})

    markers = []
    seen = set()
    for item in sorted(serials, key=lambda value: value['x']):
        serial = item['serial']
        if serial in seen:
            continue
        markers.append({'x': item['x'], 'serial': serial, 'action': ''})
        seen.add(serial)

    for idx, marker in enumerate(markers):
        left = 0 if idx == 0 else (markers[idx - 1]['x'] + marker['x']) / 2
        right = page.rect.width if idx == len(markers) - 1 else (marker['x'] + markers[idx + 1]['x']) / 2
        texts = [
            text for x0, x1, text in top_blocks
            if left <= ((x0 + x1) / 2) <= right
        ]
        for text in texts:
            action = _action_from_text(text)
            if action in ACTION_CODES:
                marker['action'] = action
                break

    return markers


def _list_type_from_text(value, current_list_type='Main'):
    list_type = current_list_type or 'Main'
    for raw_line in str(value or '').splitlines():
        compact = re.sub(r'\s+', '', raw_line)
        if 'घटक' not in compact:
            continue
        if any(term in compact for term in ('नरलयपन', 'विलोपन', 'निलोपन', 'निलयपन', 'लयपन')):
            list_type = 'Deletion'
        elif any(term in compact for term in ('ससशयधन', 'संशोधन', 'सशोधन', 'संसोधन')):
            list_type = 'Correction'
        elif any(term in compact for term in ('परररधरन', 'परिवर्धन', 'पररवर्धन', 'परीवर्धन')):
            list_type = 'Addition'
    return list_type


def _list_type_from_band(page, y0, y1, current_list_type='Main'):
    list_type = current_list_type or 'Main'
    if y1 - y0 < 3:
        return list_type
    blocks = []
    for block in page.get_text('dict').get('blocks', []):
        if block.get('type') != 0:
            continue
        bx0, by0, _bx1, by1 = block['bbox']
        if by1 <= y0 + 0.5 or by0 >= y1 - 0.5 or bx0 > page.rect.width * 0.75:
            continue
        blocks.append((by0, _pdf_block_text(block)))
    for _by0, text in sorted(blocks, key=lambda item: item[0]):
        list_type = _list_type_from_text(text, list_type)
    return list_type


def _list_type_before_page(doc, page_from):
    list_type = 'Main'
    for idx in range(max(0, min(page_from, len(doc)))):
        list_type = _list_type_from_text(doc[idx].get_text('text'), list_type)
    return list_type


def _detect_card_rows(page):
    by_y = {}
    for drawing in page.get_drawings():
        for item in drawing.get('items', []):
            if item[0] == 'l':
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) > 0.8:
                    continue
                x0, x1 = sorted((p1.x, p2.x))
                length = x1 - x0
                if 100 <= length <= 240 and p1.y > 120:
                    key = round(p1.y * 2) / 2
                    by_y.setdefault(key, []).append((x0, x1))
            elif item[0] == 're':
                rect = item[1]
                if 100 <= rect.width <= 240 and 45 <= rect.height <= 95 and rect.y0 > 120:
                    for y in (rect.y0, rect.y1):
                        key = round(y * 2) / 2
                        by_y.setdefault(key, []).append((rect.x0, rect.x1))

    y_values = _group_close_numbers(by_y.keys())
    rows = []
    for idx in range(len(y_values) - 1):
        top = y_values[idx]
        bottom = y_values[idx + 1]
        height = bottom - top
        if not 50 <= height <= 95:
            continue
        top_key = min(by_y, key=lambda value: abs(value - top))
        segments = _unique_card_segments(by_y.get(top_key, []))
        column_count = max(len(segments), _count_text_columns(page, top, bottom))
        if not column_count:
            continue
        rows.append({
            'top': top,
            'bottom': bottom,
            'count': min(column_count, 3),
            'left': min((segment[0] for segment in segments), default=0),
            'segments': segments,
        })
    return rows


def _normalize_area_line(value):
    value = str(value or '').replace('\n', ' ')
    value = re.sub(r'[_|\\[\\]{}<>~`"“”]+', ' ', value)
    value = re.sub(r'\s+', ' ', value).strip(' :-.,;')
    value = re.sub(r'\(\s+', '(', value)
    value = re.sub(r'\s+\)', ')', value)
    return value.strip()


def _looks_like_area(value):
    if not value:
        return False
    if len(re.sub(r'[\s,.:;()।-]+', '', value)) < 10:
        return False
    if len(re.findall(r'[\u0900-\u097F]', value)) < 3:
        return False
    if re.fullmatch(r'[\d\s./-]+', value):
        return False
    return not any(term in value for term in AREA_IGNORE_TERMS)


def _clean_area_text(raw_text):
    candidates = []
    for line in str(raw_text or '').splitlines():
        cleaned = _normalize_area_line(line)
        if _looks_like_area(cleaned):
            candidates.append(cleaned)
    if not candidates:
        cleaned = _normalize_area_line(raw_text)
        if _looks_like_area(cleaned):
            candidates.append(cleaned)
    return candidates[-1] if candidates else ''


def _ocr_area_block(image, bbox, dpi):
    scale = dpi / 72
    x0, y0, x1, y1 = bbox
    left = max(0, int(x0 * scale))
    top = max(0, int(y0 * scale))
    right = min(image.width, int((x1 + 4) * scale))
    bottom = min(image.height, int((y1 + 3) * scale))
    if right <= left or bottom <= top:
        return ''
    crop = image.crop((left, top, right, bottom))
    if crop.width < 30 or crop.height < 8:
        return ''
    crop = crop.resize((crop.width * 2, crop.height * 2))
    raw = pytesseract.image_to_string(crop, lang='hin', config='--psm 7')
    area = _clean_area_text(raw)
    if area:
        return area
    raw = pytesseract.image_to_string(crop, lang='eng+hin', config='--psm 7')
    return _clean_area_text(raw)


def _area_from_band(page, image, dpi, y0, y1, left_x=None):
    if y1 - y0 < 6:
        return ''
    candidates = []
    for block in page.get_text('dict').get('blocks', []):
        if block.get('type') != 0:
            continue
        bx0, by0, bx1, by1 = block['bbox']
        if by1 <= y0 + 0.5 or by0 >= y1 - 0.5:
            continue
        if bx0 > page.rect.width * 0.6:
            continue
        if left_x is not None and abs(bx0 - left_x) > 32:
            continue
        area = _ocr_area_block(image, (bx0, by0, bx1, by1), dpi)
        if area:
            candidates.append((by0, area))
    if not candidates:
        return ''
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def _apply_area_to_cells(page, image, cells, dpi, current_area=''):
    card_rows = _detect_card_rows(page)
    if not card_rows:
        for cell in cells:
            cell['area'] = current_area
        return current_area

    fallback_areas = []
    serial_area = {}
    previous_bottom = None
    for row in card_rows:
        band_top = previous_bottom if previous_bottom is not None else max(0, row['top'] - 38)
        area = _area_from_band(page, image, dpi, band_top, row['top'], row.get('left'))
        if area:
            current_area = area

        row_serials = _serials_in_row(page, row['top'], row['bottom'])
        for serial in row_serials:
            serial_area[serial] = current_area
        fallback_areas.extend([current_area] * max(row['count'], len(row_serials)))
        previous_bottom = row['bottom']

    serial_keys = sorted(serial_area, key=len, reverse=True)
    for idx, cell in enumerate(cells):
        serial = str(cell.get('serial') or '').strip()
        if serial and serial in serial_area:
            cell['area'] = serial_area[serial]
            continue

        cell_text = ' '.join(str(cell.get(key, '')) for key in (
            'crop_a', 'crop_b', 'crop_c', 'epic', 'name', 'relation_name'
        ))
        matched_area = ''
        for key in serial_keys:
            if len(key) >= 3 and re.search(rf'(?<!\d){re.escape(key)}(?!\d)', cell_text):
                matched_area = serial_area[key]
                break
        if matched_area:
            cell['area'] = matched_area
        elif idx < len(fallback_areas):
            cell['area'] = fallback_areas[idx]
        else:
            cell['area'] = current_area

    if previous_bottom is not None:
        layout_left = card_rows[0].get('left') if card_rows else None
        trailing_area = _area_from_band(page, image, dpi, previous_bottom, page.rect.height - 35, layout_left)
        if trailing_area:
            current_area = trailing_area
    return current_area


def _extract_area_only_cells(page, image, dpi, current_area='', current_list_type='Main'):
    card_rows = _detect_card_rows(page)
    cells = []
    previous_bottom = None

    for row in card_rows:
        band_top = previous_bottom if previous_bottom is not None else max(0, row['top'] - 38)
        current_list_type = _list_type_from_band(page, band_top, row['top'], current_list_type)
        area = _area_from_band(page, image, dpi, band_top, row['top'], row.get('left'))
        if area:
            current_area = area

        row_markers = _card_markers_in_row(page, row['top'], row['bottom'])
        row_count = max(row['count'], len(row_markers))
        for idx in range(row_count):
            marker = row_markers[idx] if idx < len(row_markers) else {}
            action = marker.get('action', '')
            list_type = 'Deletion' if action else current_list_type
            cells.append({
                'Serial': marker.get('serial', ''),
                'serial': marker.get('serial', ''),
                'List Type': list_type,
                'list_type': list_type,
                'Action': action,
                'action': action,
                'Area': current_area,
                'area': current_area,
                'area_only': True,
            })
        previous_bottom = row['bottom']

    if previous_bottom is not None:
        current_list_type = _list_type_from_band(page, previous_bottom, page.rect.height - 35, current_list_type)
        layout_left = card_rows[0].get('left') if card_rows else None
        trailing_area = _area_from_band(page, image, dpi, previous_bottom, page.rect.height - 35, layout_left)
        if trailing_area:
            current_area = trailing_area

    return cells, current_area, current_list_type


def _is_with_photo_file(filename):
    return 'withphoto' in _uploaded_basename(filename).replace(' ', '').lower()


def _block_text(block):
    if block.get('type') == 0:
        return _pdf_block_text(block)
    return ''


def _blocks_in_bbox(blocks, bbox, block_type=None):
    left, top, right, bottom = bbox
    matched = []
    for block in blocks:
        if block_type is not None and block.get('type') != block_type:
            continue
        x0, y0, x1, y1 = block['bbox']
        cx = (x0 + x1) / 2
        cy = (y0 + y1) / 2
        if left <= cx <= right and top <= cy <= bottom:
            matched.append(block)
    return matched


def _text_in_bbox(blocks, bbox):
    text_blocks = _blocks_in_bbox(blocks, bbox, block_type=0)
    ordered = sorted(text_blocks, key=lambda block: (block['bbox'][1], block['bbox'][0]))
    return '\n'.join(filter(None, (_block_text(block) for block in ordered)))


def _photo_left_in_card(blocks, bbox):
    image_blocks = _blocks_in_bbox(blocks, bbox, block_type=1)
    if not image_blocks:
        return None
    return min(block['bbox'][0] for block in image_blocks)


def _extract_epic(value):
    match = re.search(r'(?:[A-Z]{2,4}\d{6,}|RJ/\d{2}/\d{3}/\d{6})', str(value or ''))
    return match.group(0) if match else ''


def _normalize_ocr_text(value):
    text = str(value or '').replace('\r', '\n')
    text = re.sub(r'[|_=~`"“”<>\\[\\]{}]+', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s+', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _clean_person_value(value):
    value = _normalize_ocr_text(value)
    value = value.replace('\u200c', '').replace('\u200d', '')
    value = re.sub(r'\b(?:Photo|Available)\b', ' ', value, flags=re.IGNORECASE)
    value = re.sub(r'(?is)(?:^|\n)\s*(?:नाम|ATA|ara)\s*[:：;]\s*', ' ', value)
    value = re.sub(
        r'(?is)(?:^|\n)\s*(?:पिता|पति|माता|अन्य)\s*का\s*(?:नाम|ATA|ara)?\s*[:：]?.*$',
        ' ',
        value,
    )
    value = re.sub(r'(?is)(?:^|\n)\s*(?:मकान\s+संख्या\s*[:：]?|आयु\s*[:：]|लिंग\s*[:：]?).*$', ' ', value)
    value = value.translate(DEVANAGARI_DIGITS)
    value = re.sub(r'(?:RJ/\d{2}/\d{3}/\d{3,}|[A-Z]{2,4}\d{4,})', ' ', value)
    value = re.sub(r'\d+', ' ', value)
    value = re.sub(r'[^\u0900-\u097FA-Za-z .]+', ' ', value)
    value = re.sub(r'्\s+', '्', value)
    value = re.sub(r'\s+', ' ', value).strip(' :-.,;')
    value = re.sub(r'पुरू\s+ष', 'पुरूष', value)
    value = value.replace('पुरूषघोतम', 'पुरूषोतम')
    value = value.replace('खल्रेहा', 'स्नेहा').replace('स्रेहा', 'स्नेहा').replace('स्लेहा', 'स्नेहा')
    value = re.sub(r'(?:\s+(?:है|हे|हा|हि|ह|े|ै|ि|ी|ु|ू))+$', '', value).strip()
    value = re.sub(r'(\S+)\s+\1$', r'\1', value).strip()
    return value


def _clean_house_number(value):
    raw_value = str(value or '').translate(DEVANAGARI_DIGITS)
    value = _normalize_ocr_text(raw_value)
    value = value.splitlines()[0] if value else ''
    value = re.sub(r'(?:आयु|लिंग|नाम|पिता|पति|माता|अन्य).*$', ' ', value)
    value = re.sub(r'[|_=~`"“”<>\\[\\]{}]+', ' ', value)
    value = re.sub(r'\s+', ' ', value).strip(' :-;,')
    if value:
        return value
    return '.' if re.search(r'[.,]', raw_value) else ''


def _normalize_gender(value):
    text = str(value or '')
    compact = re.sub(r'\s+', '', text)
    if any(term in compact for term in ('पुरूष', 'पुरुष', 'पचरष')):
        return 'पुरूष'
    if 'स्त्री' in compact or 'महिला' in compact or re.search(r'(?<![A-Za-z])at(?![A-Za-z])', text, flags=re.IGNORECASE):
        return 'स्त्री'
    if re.search(r'(?<![\u0900-\u097F])सल(?![\u0900-\u097F])', text):
        return 'स्त्री'
    return ''


def _relation_from_hindi_label(label):
    if 'पति' in label or 'पनत' in label:
        return 'husband'
    if 'पिता' in label or 'नपतर' in label:
        return 'father'
    if 'माता' in label or 'मरतर' in label:
        return 'mother'
    if 'अन्य' in label:
        return 'other'
    return ''


def _parse_with_photo_ocr_text(raw_text):
    text = _normalize_ocr_text(raw_text)
    parsed = {
        'epic': _extract_epic(text),
        'name': '',
        'relation': '',
        'f_name': '',
        'relation_name': '',
        'house_number': '',
        'age': '',
        'gender': '',
        'details': text,
    }

    name_match = re.search(
        r'(?:^|\n)\s*(?:नाम|ATA|ara)\s*[:：;]\s*(.*?)(?=\n\s*(?:पिता|पति|माता|अन्य)\s*का\s*(?:नाम|ATA|ara)?\s*[:：]?|\n\s*मकान\s+संख्या\s*[:：]?|\n\s*आयु\s*[:：]?|\n\s*लिंग\s*[:：]?|$)',
        text,
        flags=re.DOTALL | re.IGNORECASE,
    )
    if name_match:
        parsed['name'] = _clean_person_value(name_match.group(1))

    relation_match = re.search(
        r'(पिता|पति|माता|अन्य)\s*का\s*(?:नाम|ATA|ara)?\s*[:：]?\s*(.*?)(?=\n\s*मकान\s+संख्या\s*[:：]?|\n\s*आयु\s*[:：]?|\n\s*लिंग\s*[:：]?|$)',
        text,
        flags=re.DOTALL | re.IGNORECASE,
    )
    if relation_match:
        parsed['relation'] = _relation_from_hindi_label(relation_match.group(1))
        parsed['f_name'] = _clean_person_value(relation_match.group(2))
        parsed['relation_name'] = parsed['f_name']

    house_match = re.search(r'मकान\s+संख्या\s*[:：]?\s*([^\n]+)', text)
    if house_match:
        parsed['house_number'] = _clean_house_number(house_match.group(1))

    normalized = text.translate(DEVANAGARI_DIGITS)
    age_match = re.search(r'आयु\s*[:：]?\s*(\d{1,3})', normalized)
    if age_match:
        parsed['age'] = age_match.group(1)

    gender_match = re.search(r'लिंग\s*[:：]?\s*([^\n]+)', text)
    parsed['gender'] = _normalize_gender(gender_match.group(1) if gender_match else text)
    return parsed


def _parse_with_photo_pdf_fallbacks(pdf_text):
    normalized = str(pdf_text or '').translate(DEVANAGARI_DIGITS)
    values = {
        'epic': _extract_epic(normalized),
        'age': '',
        'house_number': '',
        'gender': _normalize_gender(normalized),
    }
    age_match = re.search(r'(?:आजच|आयु|आयच)\s*[:：]?\s*(\d{1,3})', normalized)
    if age_match:
        values['age'] = age_match.group(1)

    # In these PDFs, the text layer often stores age, gender, and house in one
    # line like "आजच:  61 पचरष 5".
    age_line = next((line for line in normalized.splitlines() if any(term in line for term in ('आजच', 'आयु', 'आयच'))), '')
    nums = re.findall(r'\d+', age_line)
    if len(nums) >= 2:
        values['house_number'] = nums[-1]
    return values


def _with_photo_parse_score(parsed):
    score = 0
    for key in ('name', 'relation', 'f_name', 'house_number', 'age', 'gender'):
        if parsed.get(key):
            score += 2
    person_text = f"{parsed.get('name', '')} {parsed.get('f_name', '')}"
    score -= len(re.findall(r'[A-Za-z]+', person_text)) * 2
    return score


def _with_photo_needs_retry(parsed):
    if not parsed.get('name') or not parsed.get('relation') or not parsed.get('f_name'):
        return True
    return bool(re.search(r'[A-Za-z]+', f"{parsed.get('name', '')} {parsed.get('f_name', '')}"))


def _person_value_score(value):
    text = str(value or '')
    devanagari = len(re.findall(r'[\u0900-\u097F]', text))
    ascii_words = len(re.findall(r'[A-Za-z]+', text))
    label_penalty = len(re.findall(r'(?:मकान|आयु|लिंग|पिता|पति|माता|अन्य|Photo|Available)', text, flags=re.IGNORECASE))
    tokens = [token for token in re.split(r'\s+', text.strip()) if token]
    long_penalty = max(0, len(tokens) - 4) * 8 + max(0, len(text) - 32)
    noise_penalty = len(re.findall(r'[।|/=<>_~`]+|्{2,}', text)) * 4
    return devanagari - (ascii_words * 8) - (label_penalty * 10) - long_penalty - noise_penalty


def _has_latin_text(value):
    return bool(re.search(r'[A-Za-z]', str(value or '')))


def _parse_name_line_ocr_text(raw_text):
    text = _normalize_ocr_text(raw_text)
    relation_match = re.search(r'(?:पिता|पति|पाति|पॉति|पत्ति|पिला|पित्ता|माता|अन्य)\s*का', text)
    if relation_match:
        text = text[:relation_match.start()]
    label_match = re.search(r'(?:नाम|नरम|गाम|ताम)\s*[:：;ः]?\s*(.*)$', text, flags=re.DOTALL)
    if label_match:
        text = label_match.group(1)
    text = re.sub(r'(?:पिता|पति|पाति|पॉति|पत्ति|पिला|पित्ता|माता|अन्य)\s*का\s*(?:नाम|नरम|गाम)?.*$', ' ', text)
    text = re.sub(r'(?:मकान\s+संख्या|आयु|लिंग).*$', ' ', text)
    cleaned = _clean_person_value(text)
    cleaned = re.sub(r'(\S+)\s+\1$', r'\1', cleaned)
    return cleaned


def _parse_relation_line_ocr_text(raw_text):
    text = _normalize_ocr_text(raw_text)
    relation_match = re.search(
        r'(पिता|पति|पाति|पॉति|पत्ति|पिला|पित्ता|माता|अन्य)\s*का\s*(?:नाम|नरम|गाम|ताम)?\s*[:：;ः]?\s*(.*?)(?=मकान\s+संख्या|आयु|लिंग|$)',
        text,
        flags=re.DOTALL,
    )
    if not relation_match:
        return '', ''
    relation = _relation_from_hindi_label(relation_match.group(1))
    f_name = _clean_person_value(relation_match.group(2))
    return relation, f_name


def _ocr_with_photo_card(image, bbox, dpi, enhance=False):
    scale = dpi / 72
    left, top, right, bottom = bbox
    crop_box = (
        max(0, int(left * scale)),
        max(0, int(top * scale)),
        min(image.width, int(right * scale)),
        min(image.height, int(bottom * scale)),
    )
    if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
        return ''
    crop = image.crop(crop_box).convert('L')
    if enhance:
        crop = ImageEnhance.Contrast(crop).enhance(1.8).filter(ImageFilter.SHARPEN)
    crop = crop.resize((crop.width * 2, crop.height * 2))
    return pytesseract.image_to_string(crop, lang='eng+hin', config='--psm 6')


def _ocr_with_photo_line(image, bbox, dpi, lang='hin', enhance=False, scaleup=3):
    scale = dpi / 72
    left, top, right, bottom = bbox
    crop_box = (
        max(0, int(left * scale)),
        max(0, int(top * scale)),
        min(image.width, int(right * scale)),
        min(image.height, int(bottom * scale)),
    )
    if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
        return ''
    crop = image.crop(crop_box).convert('L')
    if enhance:
        crop = ImageEnhance.Contrast(crop).enhance(1.8).filter(ImageFilter.SHARPEN)
    crop = crop.resize((crop.width * scaleup, crop.height * scaleup))
    return pytesseract.image_to_string(crop, lang=lang, config='--psm 6')


def _best_with_photo_name_from_line(image, bbox, dpi, current_name=''):
    candidates = [current_name or '']
    for lang, enhance, scaleup in (
        ('hin', False, 3),
        ('eng+hin', False, 3),
        ('hin', False, 4),
        ('eng+hin', False, 4),
        ('hin', True, 3),
    ):
        raw = _ocr_with_photo_line(image, bbox, dpi, lang=lang, enhance=enhance, scaleup=scaleup)
        parsed = _parse_name_line_ocr_text(raw)
        if parsed:
            candidates.append(parsed)
    return max(candidates, key=_person_value_score).strip()


def _best_with_photo_relation_from_line(image, bbox, dpi, current_relation='', current_f_name=''):
    candidates = [(current_relation or '', current_f_name or '')]
    for lang, enhance, scaleup in (
        ('hin', False, 3),
        ('hin', False, 4),
        ('eng+hin', False, 3),
        ('hin', True, 3),
    ):
        raw = _ocr_with_photo_line(image, bbox, dpi, lang=lang, enhance=enhance, scaleup=scaleup)
        relation, f_name = _parse_relation_line_ocr_text(raw)
        if f_name:
            candidates.append((relation, f_name))
    return max(candidates, key=lambda item: _person_value_score(item[1]))


def _segment_for_marker(row, marker, idx):
    marker_x = marker.get('x')
    for left, right in row.get('segments') or []:
        if marker_x is not None and left <= marker_x <= right:
            return left, right
    segments = row.get('segments') or []
    if idx < len(segments):
        return segments[idx]
    return row.get('left', 0), row.get('left', 0) + (row.get('bottom', 0) - row.get('top', 0)) * 2.4


def _extract_with_photo_cells(page, image, dpi, current_list_type='Main'):
    card_rows = _detect_card_rows(page)
    page_blocks = page.get_text('dict').get('blocks', [])
    cells = []
    previous_bottom = None
    cell_index = 0

    for row in card_rows:
        band_top = previous_bottom if previous_bottom is not None else max(0, row['top'] - 38)
        current_list_type = _list_type_from_band(page, band_top, row['top'], current_list_type)
        markers = _card_markers_in_row(page, row['top'], row['bottom'])

        for idx, marker in enumerate(markers):
            left, right = _segment_for_marker(row, marker, idx)
            card_bbox = (left, row['top'], right, row['bottom'])
            photo_left = _photo_left_in_card(page_blocks, card_bbox)
            text_right = (photo_left - 3) if photo_left else (right - 55)
            text_right = max(left + 80, min(text_right, right - 8))
            ocr_bbox = (left, row['top'], text_right, row['bottom'])

            raw_text = _ocr_with_photo_card(image, ocr_bbox, dpi)
            parsed = _parse_with_photo_ocr_text(raw_text)
            candidates = [(parsed, raw_text)]
            if photo_left and _with_photo_needs_retry(parsed):
                retry_right = max(left + 80, min(photo_left + 15, right - 2))
                retry_raw = _ocr_with_photo_card(image, (left, row['top'], retry_right, row['bottom']), dpi)
                retry_parsed = _parse_with_photo_ocr_text(retry_raw)
                candidates.append((retry_parsed, retry_raw))
                enhanced_raw = _ocr_with_photo_card(image, ocr_bbox, dpi, enhance=True)
                enhanced_parsed = _parse_with_photo_ocr_text(enhanced_raw)
                candidates.append((enhanced_parsed, enhanced_raw))
                enhanced_retry_raw = _ocr_with_photo_card(
                    image,
                    (left, row['top'], retry_right, row['bottom']),
                    dpi,
                    enhance=True,
                )
                enhanced_retry_parsed = _parse_with_photo_ocr_text(enhanced_retry_raw)
                candidates.append((enhanced_retry_parsed, enhanced_retry_raw))
                parsed, raw_text = max(candidates, key=lambda item: _with_photo_parse_score(item[0]))

            if not parsed.get('name') or _has_latin_text(parsed.get('name')):
                name_line_bbox = (
                    left,
                    row['top'] + 13,
                    text_right,
                    min(row['bottom'], row['top'] + 42),
                )
                line_name = _best_with_photo_name_from_line(image, name_line_bbox, dpi, parsed.get('name', ''))
                if line_name and _person_value_score(line_name) > _person_value_score(parsed.get('name')):
                    parsed['name'] = line_name

            if not parsed.get('f_name') or _has_latin_text(parsed.get('f_name')):
                relation_line_bbox = (
                    left,
                    row['top'] + 27,
                    text_right,
                    min(row['bottom'], row['top'] + 51),
                )
                line_relation, line_f_name = _best_with_photo_relation_from_line(
                    image,
                    relation_line_bbox,
                    dpi,
                    parsed.get('relation', ''),
                    parsed.get('f_name', ''),
                )
                if line_f_name and _person_value_score(line_f_name) > _person_value_score(parsed.get('f_name')):
                    parsed['f_name'] = line_f_name
                    parsed['relation_name'] = line_f_name
                    if line_relation:
                        parsed['relation'] = line_relation

            pdf_text = _text_in_bbox(page_blocks, card_bbox)
            fallbacks = _parse_with_photo_pdf_fallbacks(pdf_text)

            for key in ('epic', 'age', 'house_number', 'gender'):
                if (key == 'epic' and fallbacks.get(key)) or (not parsed.get(key) and fallbacks.get(key)):
                    parsed[key] = fallbacks[key]

            action = marker.get('action', '')
            list_type = 'Deletion' if action else current_list_type
            cell_index += 1
            cells.append({
                'cell_index': cell_index,
                'row': len(cells) // 3,
                'col': idx,
                'serial': marker.get('serial', ''),
                'list_type': list_type,
                'action': action,
                'epic': parsed.get('epic', ''),
                'name': parsed.get('name', ''),
                'relation': parsed.get('relation', ''),
                'f_name': parsed.get('f_name', ''),
                'relation_name': parsed.get('relation_name', ''),
                'house_number': parsed.get('house_number', ''),
                'age': parsed.get('age', ''),
                'gender': parsed.get('gender', ''),
                'crop_a': '',
                'crop_b': '',
                'crop_c': parsed.get('details', ''),
                'details': parsed.get('details', ''),
            })

        previous_bottom = row['bottom']

    if previous_bottom is not None:
        current_list_type = _list_type_from_band(page, previous_bottom, page.rect.height - 35, current_list_type)
    return cells, current_list_type


def _fill_missing_house_numbers(cells):
    last_house = ''
    for cell in cells:
        house = str(cell.get('house_number') or '').strip()
        if house:
            last_house = house
        elif last_house:
            cell['house_number'] = last_house


def extract_voter_grid(image):
    """Replicate the APK's extraction algorithm."""
    img_w, img_h = image.size
    scale = img_w / 2479  # scale relative to reference image width

    def s(val):
        return max(1, int(val * scale))

    header_h = s(HEADER_HEIGHT)
    gap = s(START_GAP)
    card_w = s(CARD_WIDTH)
    card_h = s(CARD_HEIGHT)

    results = []

    # Process header row (cellIndex 0)
    if img_h > header_h:
        header_img = image.crop((0, 0, img_w, header_h))
        header_text = ocr_image(header_img)
        header_text = header_text.replace('\n', ' | ').strip()
    else:
        header_text = ""

    # Process data cells
    for cell_idx in range(1, MAX_CELLS + 1):
        row_num = (cell_idx - 1) // 3
        col_num = (cell_idx + 2) % 3

        from_x = gap + (card_w * col_num)
        from_y = header_h + (card_h * row_num)

        if from_x + card_w > img_w or from_y + card_h > img_h:
            break

        crop_half_w = card_w // 2
        crop_a_h = s(64)
        crop_b_h = s(64)
        crop_c_y_off = s(70)
        crop_c_w = card_w - (card_w // 3) + s(32)
        crop_c_h = card_h - s(120)

        # Crop A - left half, top portion (serial no / left data)
        try:
            crop_a = image.crop((from_x, from_y, from_x + crop_half_w, from_y + crop_a_h))
            text_a = ocr_image_confidence(crop_a)
        except Exception:
            text_a = ""

        # Crop B - right half, top portion (name / right data)
        try:
            crop_b = image.crop((from_x + crop_half_w, from_y, from_x + card_w, from_y + crop_b_h))
            text_b = ocr_image_confidence(crop_b)
        except Exception:
            text_b = ""

        # Crop C - bottom portion (EPIC / address)
        try:
            crop_c = image.crop((from_x, from_y + crop_c_y_off, from_x + crop_c_w, from_y + crop_c_y_off + crop_c_h))
            text_c = ocr_image_confidence(crop_c)
        except Exception:
            text_c = ""

        results.append({
            'cell_index': cell_idx,
            'row': row_num,
            'col': col_num,
            'crop_a': text_a,
            'crop_b': text_b,
            'crop_c': text_c,
        })

    return header_text, results


def _parse_int(value, default=0, minimum=0, maximum=None):
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        parsed = default
    parsed = max(minimum, parsed)
    if maximum is not None:
        parsed = min(maximum, parsed)
    return parsed


def _uploaded_basename(filename, fallback='upload.pdf'):
    raw_name = (filename or '').replace('\\', '/').split('/')[-1]
    safe_name = raw_name.replace('\x00', '').strip()
    if not safe_name or safe_name in ('.', '..'):
        return fallback
    return safe_name


def _file_stem(filename, fallback='voter_ocr_results'):
    base_name = _uploaded_basename(filename, fallback=f'{fallback}.pdf')
    stem, _ext = os.path.splitext(base_name)
    return stem or fallback


def _ward_part_from_filename(filename):
    match = re.search(
        r'Ward\s*No[-\s]*(\d+).*?Part\s*No[-\s]*(\d+)',
        _uploaded_basename(filename),
        flags=re.IGNORECASE,
    )
    if not match:
        return '', ''
    return match.group(1).zfill(3), match.group(2).zfill(3)


def _unique_path(folder, filename):
    safe_name = _uploaded_basename(filename, fallback='file')
    base, ext = os.path.splitext(safe_name)
    base = base or 'file'
    candidate = os.path.join(folder, f'{base}{ext}')
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f'{base}_{counter}{ext}')
        counter += 1
    return candidate


def _display_path(path):
    abs_path = os.path.abspath(path)
    if os.path.commonpath([BASE_DIR, abs_path]) == BASE_DIR:
        return os.path.relpath(abs_path, BASE_DIR)
    return abs_path


def _new_batch_id():
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    return f'batch_{timestamp}_{uuid.uuid4().hex[:8]}'


def _clean_relative_paths(relative_paths):
    if not isinstance(relative_paths, list):
        return []
    cleaned = []
    for value in relative_paths:
        text = str(value or '').replace('\\', '/').strip()
        if text:
            cleaned.append(text)
    return cleaned


def _folder_names_from_selection(source_folder_hint='', relative_paths=None):
    names = []
    for path in _clean_relative_paths(relative_paths):
        first_part = path.split('/', 1)[0].strip()
        if first_part:
            names.append(first_part)
    hint = str(source_folder_hint or '').replace('\\', '/').strip()
    if hint and hint != 'Selected folder':
        names.append(hint.split('/', 1)[0].strip())
    unique = []
    for name in names:
        if name and name not in unique:
            unique.append(name)
    return unique


def _selected_pdf_names(relative_paths=None):
    names = []
    for path in _clean_relative_paths(relative_paths):
        base = _uploaded_basename(path, fallback='')
        if base.lower().endswith('.pdf'):
            names.append(base)
    return set(names)


def _candidate_source_folders(folder_names):
    if not folder_names:
        return []
    wanted = set(folder_names)
    candidates = []
    for search_root in LOCAL_FOLDER_SEARCH_ROOTS:
        if not os.path.isdir(search_root):
            continue
        for root, dirs, _files in os.walk(search_root):
            for dirname in dirs:
                if dirname in wanted:
                    candidates.append(os.path.join(root, dirname))
    return candidates


def _resolve_selected_source_folder(source_folder_hint='', relative_paths=None):
    folder_names = _folder_names_from_selection(source_folder_hint, relative_paths)
    candidates = _candidate_source_folders(folder_names)
    if not candidates:
        return ''

    selected_files = _selected_pdf_names(relative_paths)
    if not selected_files:
        return candidates[0] if len(candidates) == 1 else ''

    best_folder = ''
    best_score = 0
    for folder in candidates:
        try:
            local_files = {
                entry.name for entry in os.scandir(folder)
                if entry.is_file() and entry.name.lower().endswith('.pdf')
            }
        except OSError:
            continue
        score = len(selected_files & local_files)
        if score > best_score:
            best_folder = folder
            best_score = score

    return best_folder if best_score else ''


def _batch_output_dir(batch_id, include_area=False, source_folder_hint='', relative_paths=None):
    if include_area:
        return app.config['AREA_EXCEL_OUTPUT_FOLDER']
    if batch_id in BATCH_OUTPUT_DIRS:
        return BATCH_OUTPUT_DIRS[batch_id]

    source_folder = _resolve_selected_source_folder(source_folder_hint, relative_paths)
    if source_folder:
        output_dir = os.path.join(os.path.dirname(source_folder), BATCH_OUTPUT_FOLDER_NAME)
    else:
        output_dir = os.path.join(app.config['EXCEL_OUTPUT_FOLDER'], batch_id)

    BATCH_OUTPUT_DIRS[batch_id] = output_dir
    return output_dir


def _batch_dirs(batch_id=None, include_area=False, source_folder_hint='', relative_paths=None):
    batch_id = batch_id or _new_batch_id()
    safe_batch_id = secure_filename(batch_id)
    if safe_batch_id != batch_id or not batch_id.startswith('batch_'):
        raise ValueError('Invalid batch id')

    staging_dir = os.path.join(app.config['UPLOAD_FOLDER'], batch_id)
    output_dir = _batch_output_dir(
        batch_id,
        include_area=include_area,
        source_folder_hint=source_folder_hint,
        relative_paths=relative_paths,
    )
    processed_dir = os.path.join(app.config['PROCESSED_PDF_FOLDER'], batch_id)
    for folder in (staging_dir, output_dir, processed_dir):
        os.makedirs(folder, exist_ok=True)
    return batch_id, staging_dir, output_dir, processed_dir


def _rows_from_cells(cells):
    if cells and all(c.get('area_only') for c in cells):
        return _area_only_rows_from_cells(cells)
    if any(c.get('serial') or c.get('name') or c.get('ward_no') for c in cells):
        return _full_rows_from_cells(cells)

    include_area = any('area' in c for c in cells)
    rows = []
    for c in cells:
        row = {
            'Page': c.get('page', 0),
        }
        if include_area:
            row['Area'] = c.get('area', '')
        row.update({
            'Header': c.get('header', ''),
            'Details': c.get('crop_c', ''),
        })
        rows.append(row)
    return rows


def _full_rows_from_cells(cells):
    rows = []
    for c in cells:
        rows.append({
            'WARD_NO': c.get('ward_no', ''),
            'PART_NO': c.get('part_no', ''),
            'Page': c.get('page', 0),
            'List Type': c.get('list_type', ''),
            'Action': c.get('action', ''),
            'Serial': c.get('serial') or c.get('cell_index', ''),
            'EPIC': c.get('epic', ''),
            'Name': c.get('name', ''),
            'Relation': c.get('relation', ''),
            'F_NAME': c.get('f_name') or c.get('relation_name', ''),
            'House Number': c.get('house_number', ''),
            'Age': c.get('age', ''),
            'Gender': c.get('gender', ''),
        })
    return rows


def _area_only_rows_from_cells(cells):
    rows = []
    serial_indexes = {}
    for c in cells:
        serial = str(c.get('Serial') or c.get('serial') or '').strip()
        if not serial:
            continue
        list_type = c.get('List Type') or c.get('list_type', '')
        action = c.get('Action') or c.get('action', '')
        row = {
            'Serial': serial,
            'List Type': list_type,
            'Action': action,
            'Area': c.get('Area') or c.get('area', ''),
        }
        if serial in serial_indexes:
            rows[serial_indexes[serial]] = row
        else:
            serial_indexes[serial] = len(rows)
            rows.append(row)
    return rows


def _dataframe_from_cells(cells, area_only=False):
    if area_only or (cells and all(c.get('area_only') for c in cells)):
        return pd.DataFrame(_area_only_rows_from_cells(cells), columns=AREA_ONLY_EXPORT_COLUMNS)
    if any(c.get('serial') or c.get('name') or c.get('ward_no') for c in cells):
        return pd.DataFrame(_full_rows_from_cells(cells), columns=FULL_EXPORT_COLUMNS)
    return pd.DataFrame(_rows_from_cells(cells))


def _write_excel(cells, output_path, area_only=False):
    df = _dataframe_from_cells(cells, area_only=area_only)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='OCR Results')


def _save_processed_excel_and_pdf(staged_path, original_name, output_dir, processed_dir, result):
    excel_path = _unique_path(output_dir, f'{_file_stem(original_name)}.xlsx')
    _write_excel(result['cells'], excel_path, area_only=result.get('area_only', False))

    processed_path = _unique_path(processed_dir, original_name)
    shutil.move(staged_path, processed_path)

    return {
        'filename': original_name,
        'excel_path': _display_path(excel_path),
        'processed_pdf_path': _display_path(processed_path),
        'cells': len(_area_only_rows_from_cells(result['cells'])) if result.get('area_only') else len(result['cells']),
        'pages': result['page_to'] - result['page_from'] + 1,
    }


def _iter_extract_pdf(filepath, page_from, page_to, dpi, include_area=False):
    try:
        doc = fitz.open(filepath)
    except Exception as exc:
        yield {'type': 'error', 'message': f'Could not open PDF: {exc}'}
        return

    try:
        total_pages = len(doc)
        if total_pages == 0:
            yield {'type': 'error', 'message': 'PDF has no pages'}
            return

        page_to_actual = page_to
        if page_to_actual == 0 or page_to_actual >= total_pages:
            page_to_actual = total_pages - 1

        if page_from > page_to_actual:
            yield {'type': 'error', 'message': 'From page cannot be after To page'}
            return
        if page_from >= total_pages:
            yield {'type': 'error', 'message': 'From page exceeds PDF page count'}
            return

        total = page_to_actual - page_from + 1
        all_cells = []
        all_headers = {}
        current_area = ''
        current_list_type = _list_type_before_page(doc, page_from)
        ward_no, part_no = _ward_part_from_filename(os.path.basename(filepath))
        is_with_photo = _is_with_photo_file(os.path.basename(filepath))
        start_time = time.time()
        page_times = []

        for i, page_num in enumerate(range(page_from, page_to_actual + 1)):
            t0 = time.time()

            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = doc[page_num].get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            if include_area:
                header = ''
                cells, current_area, current_list_type = _extract_area_only_cells(
                    doc[page_num],
                    img,
                    dpi,
                    current_area,
                    current_list_type,
                )
            else:
                if is_with_photo:
                    header = ''
                    cells, current_list_type = _extract_with_photo_cells(
                        doc[page_num],
                        img,
                        dpi,
                        current_list_type,
                    )
                else:
                    header, cells = extract_voter_grid(img)
            all_headers[str(page_num)] = header
            for c in cells:
                if not include_area:
                    c['page'] = page_num
                    c['header'] = header
                    c['ward_no'] = ward_no
                    c['part_no'] = part_no
                all_cells.append(c)

            t1 = time.time()
            page_times.append(t1 - t0)
            elapsed = t1 - start_time
            avg = sum(page_times) / len(page_times)
            eta = avg * (total - (i + 1))

            yield {
                'type': 'progress',
                'current': i + 1,
                'total': total,
                'elapsed': round(elapsed, 1),
                'eta': round(eta, 1),
            }

        if not include_area and is_with_photo:
            _fill_missing_house_numbers(all_cells)

        yield {
            'type': 'complete',
            'cells': all_cells,
            'headers': all_headers,
            'page_from': page_from,
            'page_to': page_to_actual,
            'total_pages': total_pages,
            'include_area': include_area,
            'area_only': include_area,
        }
    finally:
        doc.close()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/info', methods=['POST'])
def info():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file'}), 400
    fname = _uploaded_basename(file.filename)
    fpath = os.path.join(app.config['UPLOAD_FOLDER'], fname)
    file.save(fpath)
    doc = fitz.open(fpath)
    total = len(doc)
    doc.close()
    return jsonify({'total_pages': total, 'filename': fname})


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return _sse(json.dumps({'type': 'error', 'message': 'No file uploaded'}))

    file = request.files['file']
    if not file.filename:
        return _sse(json.dumps({'type': 'error', 'message': 'No file selected'}))

    filename = _uploaded_basename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    page_from = _parse_int(request.form.get('page_from'), default=0)
    page_to = _parse_int(request.form.get('page_to'), default=0)
    dpi = _parse_int(request.form.get('dpi'), default=200, minimum=100, maximum=600)
    include_area = _parse_bool(request.form.get('include_area'))

    def generate():
        result = None
        for event in _iter_extract_pdf(filepath, page_from, page_to, dpi, include_area=include_area):
            if event['type'] == 'progress':
                yield _sse(json.dumps(event))
            elif event['type'] == 'error':
                yield _sse(json.dumps(event))
                return
            elif event['type'] == 'complete':
                result = event

        if result is None:
            yield _sse(json.dumps({'type': 'error', 'message': 'Processing did not finish'}))
            return

        # preview
        docpre = fitz.open(filepath)
        matpre = fitz.Matrix(dpi / 72, dpi / 72)
        pixpre = docpre[result['page_from']].get_pixmap(matrix=matpre)
        imgpre = Image.frombytes("RGB", [pixpre.width, pixpre.height], pixpre.samples)
        docpre.close()

        thumb = imgpre.copy()
        thumb.thumbnail((600, 800))
        buf = io.BytesIO()
        thumb.save(buf, format='PNG')
        thumb_b64 = base64.b64encode(buf.getvalue()).decode()

        yield _sse(json.dumps({
            'type': 'result',
            'cells': result['cells'],
            'headers': result['headers'],
            'preview': f'data:image/png;base64,{thumb_b64}',
            'page_from': result['page_from'],
            'page_to': result['page_to'],
            'total_pages': result['total_pages'],
            'include_area': result.get('include_area', False),
            'area_only': result.get('area_only', False),
            'filename': filename,
            'image_size': f'{imgpre.width}x{imgpre.height}'
        }))

    return Response(generate(), mimetype='text/event-stream')


@app.route('/batch-start', methods=['POST'])
def batch_start():
    data = request.get_json(silent=True) or {}
    include_area = _parse_bool(data.get('include_area'))
    source_folder_hint = data.get('source_folder_hint', '')
    relative_paths = data.get('relative_paths', [])
    try:
        batch_id, _staging_dir, output_dir, processed_dir = _batch_dirs(
            include_area=include_area,
            source_folder_hint=source_folder_hint,
            relative_paths=relative_paths,
        )
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400

    return jsonify({
        'batch_id': batch_id,
        'output_folder': _display_path(output_dir),
        'processed_folder': _display_path(processed_dir),
    })


@app.route('/batch-process-file', methods=['POST'])
def batch_process_file():
    file = request.files.get('file')
    if not file or not file.filename or not file.filename.lower().endswith('.pdf'):
        return Response(_sse(json.dumps({
            'type': 'file_error',
            'filename': '',
            'error': 'No PDF file selected'
        })), mimetype='text/event-stream')

    batch_id = request.form.get('batch_id')
    if not batch_id:
        return Response(_sse(json.dumps({
            'type': 'file_error',
            'filename': _uploaded_basename(file.filename),
            'error': 'Missing batch id',
        })), mimetype='text/event-stream')

    original_name = _uploaded_basename(file.filename)
    page_from = _parse_int(request.form.get('page_from'), default=0)
    page_to = _parse_int(request.form.get('page_to'), default=0)
    dpi = _parse_int(request.form.get('dpi'), default=200, minimum=100, maximum=600)
    include_area = _parse_bool(request.form.get('include_area'))
    file_index = _parse_int(request.form.get('file_index'), default=1, minimum=1)
    total_files = _parse_int(request.form.get('total_files'), default=1, minimum=1)

    try:
        _batch_id, staging_dir, output_dir, processed_dir = _batch_dirs(batch_id, include_area=include_area)
    except ValueError as exc:
        return Response(_sse(json.dumps({
            'type': 'file_error',
            'filename': original_name,
            'error': str(exc),
        })), mimetype='text/event-stream')

    staged_path = _unique_path(staging_dir, original_name)

    save_error = None
    try:
        file.save(staged_path)
    except Exception as exc:
        save_error = str(exc)

    def generate():
        yield _sse(json.dumps({
            'type': 'file_start',
            'file_index': file_index,
            'total_files': total_files,
            'filename': original_name,
        }))

        if save_error:
            yield _sse(json.dumps({
                'type': 'file_error',
                'filename': original_name,
                'error': save_error,
            }))
            return

        try:
            result = None
            for event in _iter_extract_pdf(staged_path, page_from, page_to, dpi, include_area=include_area):
                if event['type'] == 'progress':
                    event.update({
                        'type': 'file_progress',
                        'file_index': file_index,
                        'total_files': total_files,
                        'filename': original_name,
                    })
                    yield _sse(json.dumps(event))
                elif event['type'] == 'error':
                    raise RuntimeError(event['message'])
                elif event['type'] == 'complete':
                    result = event

            if result is None:
                raise RuntimeError('Processing did not finish')

            success = _save_processed_excel_and_pdf(
                staged_path,
                original_name,
                output_dir,
                processed_dir,
                result,
            )
            yield _sse(json.dumps({'type': 'file_done', **success}))
        except Exception as exc:
            yield _sse(json.dumps({
                'type': 'file_error',
                'filename': original_name,
                'error': str(exc),
            }))

    return Response(generate(), mimetype='text/event-stream')


@app.route('/batch-process', methods=['POST'])
def batch_process():
    files = [
        file for file in request.files.getlist('files')
        if file and file.filename and file.filename.lower().endswith('.pdf')
    ]
    if not files:
        return Response(_sse(json.dumps({
            'type': 'error',
            'message': 'No PDF files selected'
        })), mimetype='text/event-stream')

    page_from = _parse_int(request.form.get('page_from'), default=0)
    page_to = _parse_int(request.form.get('page_to'), default=0)
    dpi = _parse_int(request.form.get('dpi'), default=200, minimum=100, maximum=600)
    include_area = _parse_bool(request.form.get('include_area'))

    batch_id, staging_dir, output_dir, processed_dir = _batch_dirs(include_area=include_area)

    staged_files = []
    for file_index, file in enumerate(files, start=1):
        original_name = _uploaded_basename(file.filename)
        staged_path = _unique_path(staging_dir, original_name)
        try:
            file.save(staged_path)
            staged_files.append({
                'file_index': file_index,
                'filename': original_name,
                'staged_path': staged_path,
                'error': None,
            })
        except Exception as exc:
            staged_files.append({
                'file_index': file_index,
                'filename': original_name,
                'staged_path': None,
                'error': str(exc),
            })

    def generate():
        successes = []
        failures = []

        yield _sse(json.dumps({
            'type': 'batch_start',
            'total_files': len(files),
            'output_folder': _display_path(output_dir),
            'processed_folder': _display_path(processed_dir),
        }))

        for item in staged_files:
            file_index = item['file_index']
            original_name = item['filename']
            staged_path = item['staged_path']

            yield _sse(json.dumps({
                'type': 'file_start',
                'file_index': file_index,
                'total_files': len(files),
                'filename': original_name,
            }))

            try:
                if item['error']:
                    raise RuntimeError(item['error'])

                result = None
                for event in _iter_extract_pdf(staged_path, page_from, page_to, dpi, include_area=include_area):
                    if event['type'] == 'progress':
                        event.update({
                            'type': 'file_progress',
                            'file_index': file_index,
                            'total_files': len(files),
                            'filename': original_name,
                        })
                        yield _sse(json.dumps(event))
                    elif event['type'] == 'error':
                        raise RuntimeError(event['message'])
                    elif event['type'] == 'complete':
                        result = event

                if result is None:
                    raise RuntimeError('Processing did not finish')

                success = _save_processed_excel_and_pdf(
                    staged_path,
                    original_name,
                    output_dir,
                    processed_dir,
                    result,
                )
                successes.append(success)
                yield _sse(json.dumps({'type': 'file_done', **success}))
            except Exception as exc:
                failure = {
                    'filename': original_name,
                    'error': str(exc),
                }
                failures.append(failure)
                yield _sse(json.dumps({'type': 'file_error', **failure}))

        yield _sse(json.dumps({
            'type': 'batch_result',
            'success_count': len(successes),
            'failure_count': len(failures),
            'successes': successes,
            'failures': failures,
            'output_folder': _display_path(output_dir),
            'processed_folder': _display_path(processed_dir),
        }))

    return Response(generate(), mimetype='text/event-stream')


def _sse(data):
    return f'data: {data}\n\n'


@app.route('/download/<fmt>', methods=['POST'])
def download(fmt):
    data = request.get_json()
    if not data or 'cells' not in data:
        return jsonify({'error': 'No data'}), 400

    df = _dataframe_from_cells(data['cells'], area_only=_parse_bool(data.get('area_only')))
    source_filename = data.get('filename') or data.get('source_filename') or 'voter_ocr_results.pdf'

    if fmt == 'csv':
        buf = io.BytesIO()
        df.to_csv(buf, index=False)
        buf.seek(0)
        return send_file(buf, mimetype='text/csv', as_attachment=True,
                         download_name=f'{_file_stem(source_filename)}.csv')
    elif fmt == 'xlsx':
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='OCR Results')
        buf.seek(0)
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=f'{_file_stem(source_filename)}.xlsx')

    return jsonify({'error': 'Invalid format'}), 400


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True, threaded=True)
