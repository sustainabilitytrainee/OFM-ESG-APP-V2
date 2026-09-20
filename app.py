from flask import Flask, request, render_template, send_file, jsonify
import openpyxl, os, io, tempfile, json, re
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

OMAN_GRID_EF = 0.561
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

MONTH_MAP = {
    "january":"Jan","february":"Feb","march":"March","april":"April",
    "may":"May","june":"Jun","july":"Jul","august":"Aug",
    "september":"Sep","october":"Oct","november":"Nov","december":"Dec",
}

def extract_bill(pdf_bytes):
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel("gemini-1.5-flash")

    prompt = """You are extracting data from an Oman electricity bill (Nama utility company).
Extract these four fields:
1. account_no: The electricity account number (format: one letter followed by digits, e.g. T04556, T11189)
2. month: The billing month name in English (e.g. January, February, June)
3. year: The billing year as a 4-digit number (e.g. 2024, 2026)
4. kwh: The total units consumed in kWh as a plain integer (e.g. 287856). Labeled "Units Consumed (KWH)" on the bill.

Return ONLY a JSON object with these exact keys: account_no, month, year, kwh
Example: {"account_no": "T04556", "month": "June", "year": 2026, "kwh": 287856}"""

    try:
        response = model.generate_content([
            {"mime_type": "application/pdf", "data": pdf_bytes},
            prompt
        ])
        raw = response.text.strip()
        raw = re.sub(r"```json\s*|\s*```", "", raw).strip()
        data = json.loads(raw)
        month_name = data.get("month", "").lower()
        month_col = MONTH_MAP.get(month_name)
        return {
            "account_no": data.get("account_no"),
            "month": month_col,
            "month_label": f"{data.get('month')} {data.get('year')}",
            "year": data.get("year"),
            "kwh": int(data.get("kwh",
