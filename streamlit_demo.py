import streamlit as st
import pandas as pd
import os
import pdfplumber
import textwrap
import json
import google.generativeai as genai
from io import BytesIO

# ---------------- CONFIG -----------------
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel("models/gemini-flash-latest")

# -------------- HELPERS ------------------
def extract_text(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            content = page.extract_text()
            if content:
                text += "\n" + content
    return text

def chunk(text, limit=5000):
    cleaned = " ".join(text.split())
    return textwrap.wrap(cleaned, limit)

def ask_model(txt):
    prompt = f"""
    Extract key:value pairs from text below.
    Rules:
    • No paraphrasing
    • Values must be exact
    • Provide a 'context' per row
    • If not identifiable, return as {{'key':'raw_text','value':txt,'context':txt}}

    Return only JSON list.
    Text: \"\"\"{txt}\"\"\"
    """
    res = model.generate_content(prompt)
    return res.text

def safe_json(data):
    s, e = data.find("["), data.rfind("]")
    if s == -1 or e == -1: return []
    try: return json.loads(data[s:e+1])
    except: return []

# ---- CLEAN COMMENTS + DUPLICATES ----
def clean_rows(records):
    unique_rows, seen_rows = [], set()
    for r in records:
        t = (r.get("key"), r.get("value"), r.get("context"))
        if t not in seen_rows:
            seen_rows.add(t)
            unique_rows.append(r)

    final, seen_context = [], set()
    for r in unique_rows:
        ctx = r.get("context", "")
        if ctx not in seen_context:
            seen_context.add(ctx)
            final.append(r)

    return final

# ---------------- UI ----------------
st.title("📄 AI Document → Excel Extractor")
st.write("Upload a PDF to convert structured key:value data automatically.")

uploaded = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded:
    st.info("Processing... please wait ⏳")

    text = extract_text(uploaded)
    parts = chunk(text)

    result = []
    for p in parts:
        out = ask_model(p)
        result.extend(safe_json(out))

    cleaned = clean_rows(result)
    df = pd.DataFrame(cleaned)
    df.index += 1
    df = df.rename(columns={"key": "Key", "value": "Value", "context": "Comments"})

    # ----- FORMAT EXCEL -----
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index_label="#", sheet_name="Output")

        # Format
        workbook = writer.book
        sheet = writer.sheets["Output"]

        # Column widths
        sheet.column_dimensions['A'].width = 4
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['C'].width = 35
        sheet.column_dimensions['D'].width = 100

        # Bold headers
        for cell in sheet["1:1"]:
            cell.font = cell.font.copy(bold=True)

    st.success("🎉 Excel Generated Successfully!")

    st.dataframe(df)

    st.download_button("⬇ Download Excel File", data=buffer.getvalue(), file_name="Output.xlsx")

