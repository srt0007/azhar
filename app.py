from flask import Flask, render_template, request, jsonify, send_file
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024


def extract_data_from_pdf(pdf_bytes, filename=''):
    results = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
            if not text:
                continue
            lines = [l.strip() for l in text.split('\n')]

            # Find tracking number on this page
            tracking_no = None
            for line in lines:
                m = re.search(r'\b(1Z[A-Z0-9]{16})\b', line, re.IGNORECASE)
                if m:
                    tracking_no = m.group(1).upper()
                    break
            if not tracking_no:
                continue

            # Parse net charges section
            in_charges = False
            for line in lines:
                if not line:
                    continue
                if re.search(r'\bNet\s*Charges\b', line, re.IGNORECASE):
                    in_charges = True
                    continue
                if in_charges and re.match(r'^(Total Charges|Date\s|Page \d|Tracking)', line, re.IGNORECASE):
                    in_charges = False
                if not in_charges:
                    continue
                if line.lower() in ['description', 'net charges']:
                    continue

                # Match "Some Description   123.45"
                m = re.match(r'^(.+?)\s+([\d,]+\.\d{2})\s*$', line)
                if m:
                    desc = m.group(1).strip()
                    skip = ['description', 'shipper', 'consignee', 'payor', 'comments', 'reference']
                    if any(s in desc.lower() for s in skip):
                        continue
                    try:
                        amount = float(m.group(2).replace(',', ''))
                        if amount > 0:
                            results.append({
                                'tracking_no': tracking_no,
                                'description': desc,
                                'net_charges': round(amount, 2)
                            })
                    except ValueError:
                        pass

            # Fallback: known charge keywords
            if not any(r['tracking_no'] == tracking_no for r in results):
                known = ['Duty Amount', 'Value Added Tax', 'Import Tax', 'Customs Duty',
                         'Fuel Surcharge', 'Peak Surcharge', 'Residential Surcharge',
                         'Disbursement Fee', 'Handling Fee', 'Transportation Charge']
                for charge in known:
                    m = re.search(rf'{re.escape(charge)}\s+([\d,]+\.\d{{2}})', text, re.IGNORECASE)
                    if m:
                        try:
                            amount = float(m.group(1).replace(',', ''))
                            if amount > 0:
                                results.append({
                                    'tracking_no': tracking_no,
                                    'description': charge,
                                    'net_charges': round(amount, 2)
                                })
                        except ValueError:
                            pass

    # Deduplicate
    seen = set()
    unique = []
    for r in results:
        key = (r['tracking_no'], r['description'].lower(), r['net_charges'])
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files', 'data': [], 'total': 0}), 400

    all_data, errors = [], []
    for f in files:
        if not f.filename.lower().endswith('.pdf'):
            errors.append(f'{f.filename}: Not a PDF')
            continue
        try:
            data = extract_data_from_pdf(f.read(), f.filename)
            all_data.extend(data)
        except Exception as e:
            errors.append(f'{f.filename}: {str(e)}')

    return jsonify({'data': all_data, 'total': len(all_data), 'errors': errors})


@app.route('/download', methods=['POST'])
def download():
    data = request.json.get('data', [])
    if not data:
        return jsonify({'error': 'No data'}), 400

    # Pivot: one row per tracking number, one column per description
    descs = list(dict.fromkeys(r['description'] for r in data))  # ordered unique
    tracking_order, pivot = [], {}
    for r in data:
        tn = r['tracking_no']
        if tn not in pivot:
            pivot[tn] = {}
            tracking_order.append(tn)
        pivot[tn][r['description']] = r['net_charges']
    trackings = list(dict.fromkeys(tracking_order))

    wb = Workbook()
    ws = wb.active
    ws.title = "UPS Invoice Data"

    hfill = PatternFill(start_color="351C75", end_color="351C75", fill_type="solid")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    tfill = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")
    afill = PatternFill(start_color="F3F0FF", end_color="F3F0FF", fill_type="solid")
    zero_font = Font(color="D1D5DB", size=10)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    right = Alignment(horizontal='right', vertical='center')
    center = Alignment(horizontal='center', vertical='center')

    # Header row: Tracking No. | <desc1> | <desc2> | ... | Total
    headers = ['Tracking No.'] + descs + ['Total']
    ws.row_dimensions[1].height = 32
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hfont
        c.fill = hfill
        c.alignment = center
        c.border = border

    # Data rows
    for i, tn in enumerate(trackings, 2):
        ws.row_dimensions[i].height = 20
        ws.cell(row=i, column=1, value=tn).border = border
        row_total = 0
        for j, desc in enumerate(descs, 2):
            val = pivot[tn].get(desc, 0)
            row_total += val
            c = ws.cell(row=i, column=j, value=val)
            c.number_format = '#,##0.00'
            c.alignment = right
            c.border = border
            if val == 0:
                c.font = zero_font
            if i % 2 == 0:
                c.fill = afill
        # Row total
        tc = ws.cell(row=i, column=len(headers), value=row_total)
        tc.number_format = '#,##0.00'
        tc.alignment = right
        tc.border = border
        tc.font = Font(bold=True, size=10)
        if i % 2 == 0:
            tc.fill = afill
        if i % 2 == 0:
            ws.cell(row=i, column=1).fill = afill

    # Totals row
    total_row = len(trackings) + 2
    ws.row_dimensions[total_row].height = 24
    tc = ws.cell(row=total_row, column=1, value='Total')
    tc.font = Font(bold=True, size=11)
    tc.fill = tfill
    tc.border = border
    grand = 0
    for j, desc in enumerate(descs, 2):
        col_total = sum(pivot[tn].get(desc, 0) for tn in trackings)
        grand += col_total
        c = ws.cell(row=total_row, column=j, value=col_total)
        c.number_format = '#,##0.00'
        c.alignment = right
        c.font = Font(bold=True, size=10)
        c.fill = tfill
        c.border = border
    gc = ws.cell(row=total_row, column=len(headers), value=grand)
    gc.number_format = '#,##0.00'
    gc.alignment = right
    gc.font = Font(bold=True, size=11)
    gc.fill = tfill
    gc.border = border

    # Column widths
    ws.column_dimensions[get_column_letter(1)].width = 26
    for j in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    ws.freeze_panes = 'B2'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='ups_invoice_data.xlsx')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=False)
