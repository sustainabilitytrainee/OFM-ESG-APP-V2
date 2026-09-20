from flask import Flask, request, render_template, send_file, jsonify
import pdfplumber, openpyxl, os, io, tempfile, json, re
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

OMAN_GRID_EF = 0.561  # kg CO2e/kWh
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

MONTH_MAP = {
    "january":"Jan","february":"Feb","march":"March","april":"April",
    "may":"May","june":"Jun","july":"Jul","august":"Aug",
    "september":"Sep","october":"Oct","november":"Nov","december":"Dec",
}

def extract_bill(pdf_bytes):
    # Extract text from PDF
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # Use Gemini to extract the fields
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel("gemini-2.0-flash")

    prompt = f"""You are extracting data from an Oman electricity bill (Nama utility company).
Extract these three fields from the bill text below:
1. account_no: The electricity account number (format: one letter followed by digits, e.g. T04556, T11189)
2. month: The billing month name in English (e.g. January, February, June)
3. year: The billing year as a 4-digit number (e.g. 2024, 2026)
4. kwh: The total units consumed in kWh as a plain integer (e.g. 287856). This is labeled "Units Consumed (KWH)" on the bill.

Return ONLY a JSON object with these exact keys: account_no, month, year, kwh
Example: {{"account_no": "T04556", "month": "June", "year": 2026, "kwh": 287856}}

Bill text:
{text[:3000]}"""

    try:
        response = model.generate_content(prompt)
        raw = response.text.strip()
        # Strip markdown code blocks if present
        raw = re.sub(r"```json\s*|\s*```", "", raw).strip()
        data = json.loads(raw)
        month_name = data.get("month", "").lower()
        month_col = MONTH_MAP.get(month_name)
        return {
            "account_no": data.get("account_no"),
            "month": month_col,
            "month_label": f"{data.get('month')} {data.get('year')}",
            "year": data.get("year"),
            "kwh": int(data.get("kwh", 0)) if data.get("kwh") else None,
        }
    except Exception as e:
        return {
            "account_no": None, "month": None,
            "month_label": "Unknown", "year": None, "kwh": None,
            "gemini_error": str(e)
        }

def update_excel(excel_bytes, records):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))

    # Find electricity sheet
    sheet_name = None
    for name in wb.sheetnames:
        if "electricity" in name.lower() and "consumption" in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        raise ValueError(f"Could not find electricity sheet. Sheets: {wb.sheetnames}")
    ws = wb[sheet_name]

    # Build month column map
    month_col_map = {}
    account_col_idx = 4
    for row in ws.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if cell.value and str(cell.value).strip() in [v.strip() for v in MONTH_MAP.values()]:
                month_col_map[str(cell.value).strip()] = cell.column

    # Scope 2 sheet
    s2_name = "Scope 2 (Electricity)"
    if s2_name not in wb.sheetnames:
        ws2 = wb.create_sheet(s2_name)
        ws2.append(["Scope 2 — Purchased Electricity"])
        ws2.append([f"Emission factor: {OMAN_GRID_EF} kg CO2e/kWh (Oman grid, OPWP 2023)"])
        ws2.append([])
        ws2.append(["Account No", "Location", "Type", "Period", "Month", "kWh", "tCO2e"])
    else:
        ws2 = wb[s2_name]

    results = []
    for rec in records:
        account_no = rec.get("account_no")
        month_col_header = rec.get("month")
        kwh = rec.get("kwh")

        if not account_no or not month_col_header or not kwh:
            msg = rec.get("gemini_error", "Could not extract all fields")
            results.append({**rec, "status": "error", "message": msg})
            continue

        month_col_idx = month_col_map.get(month_col_header.strip())
        if not month_col_idx:
            results.append({**rec, "status": "error", "message": f"Month '{month_col_header}' not found in sheet"})
            continue

        # Find account row
        target_row = None
        location = ""
        unit_type = ""
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            if row[account_col_idx - 1].value and str(row[account_col_idx - 1].value).strip() == account_no:
                target_row = row[0].row
                location = row[2].value or ""
                unit_type = row[4].value or ""
                break

        if not target_row:
            results.append({**rec, "status": "error", "message": f"Account {account_no} not found in Excel"})
            continue

        ws.cell(row=target_row, column=month_col_idx).value = kwh
        row_total = sum(ws.cell(row=target_row, column=c).value or 0 for c in range(7, 19))
        ws.cell(row=target_row, column=19).value = row_total

        tco2e = round(kwh * OMAN_GRID_EF / 1000, 4)
        ws2.append([account_no, location, unit_type, rec["month_label"], month_col_header.strip(), kwh, tco2e])

        results.append({
            **rec,
            "status": "ok",
            "location": location,
            "tco2e": tco2e,
        })

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, results

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    bills = request.files.getlist("bills")
    excel = request.files.get("excel")

    if not excel or not bills:
        return jsonify({"error": "Upload both the Excel file and at least one bill PDF"}), 400

    if not GEMINI_API_KEY:
        return jsonify({"error": "Gemini API key not configured on server"}), 500

    excel_bytes = excel.read()
    records = []
    for bill in bills:
        pdf_bytes = bill.read()
        rec = extract_bill(pdf_bytes)
        rec["filename"] = bill.filename
        records.append(rec)

    try:
        updated_excel, results = update_excel(excel_bytes, records)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir="/tmp")
    tmp.write(updated_excel.read())
    tmp.flush()
    app._last_excel = tmp.name

    return jsonify({"results": results})

@app.route("/download")
def download():
    path = getattr(app, "_last_excel", None)
    if not path or not os.path.exists(path):
        return "No file ready", 404
    return send_file(path, as_attachment=True, download_name="Environmental_Metrics_Updated.xlsx")

if __name__ == "__main__":
    app.run(debug=True, port=5000)
