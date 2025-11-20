import os
import json
import textwrap
import pdfplumber
import pandas as pd
import google.generativeai as genai

# CONFIGURATION
PDF_PATH = "data/Data Input.pdf"
OUTPUT_EXCEL = "Output.xlsx"

# SET GEMINI API KEY
if not os.getenv("GEMINI_API_KEY"):
    raise Exception("⚠ Please set GEMINI_API_KEY environment variable!")

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Choose correct model shown available in your account
model = genai.GenerativeModel("models/gemini-flash-latest")

# READ PDF

def read_pdf_text(path):
    full_text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            content = page.extract_text()
            if content:
                full_text += "\n" + content
    return full_text

# SPLIT LARGE TEXT

def split_chunks(text, limit=6000):
    cleaned = " ".join(text.split())
    return textwrap.wrap(cleaned, limit)

# ASK GEMINI TO EXTRACT DATA

def extract_ai(text):
    prompt = f"""
    You are a document parser. Extract ALL possible key:value information 
    from the given text strictly in original wording.

    RULES:
    • Do NOT summarize or rephrase any text.
    • Keys must be inferred logically (like Date of Birth, Salary, Degree, Organization, Address etc.)
    • Values MUST stay exact as they appear.
    • Include a `context` for each pair, containing the FULL sentence/paragraph.
    • If a part of text cannot be turned into a key:value pair, output it as:
      {{"key":"raw_text", "value":"<text>", "context":"<text>"}}

    Return ONLY valid JSON list like:
    [
      {{"key":"...", "value":"...", "context":"..."}}
    ]

    TEXT:
    \"\"\"{text}\"\"\"
    """
    response = model.generate_content(prompt)
    return response.text

# SAFE JSON PARSING

def safe_parse(json_text):
    start = json_text.find("[")
    end = json_text.rfind("]")

    if start == -1 or end == -1:
        return [{"key": "raw_text", "value": json_text, "context": json_text}]
    try:
        return json.loads(json_text[start:end + 1])
    except:
        return [{"key": "raw_text", "value": json_text, "context": json_text}]


# MAIN PIPELINE

if __name__ == "__main__":
    print("Reading PDF...")
    pdf_text = read_pdf_text(PDF_PATH)

    print("Splitting text into chunks...")
    parts = split_chunks(pdf_text)

    final = []

    print("Extracting using Gemini Model...")
    for i, chunk in enumerate(parts, 1):
        print(f"• Processing chunk {i}/{len(parts)}...")
        output = extract_ai(chunk)
        data = safe_parse(output)
        final.extend(data)

    # REMOVE DUPLICATE ROWS (Same Key/Value/Context)

    cleaned = []
    seen = set()

    for row in final:
        triple = (row.get("key"), row.get("value"), row.get("context"))
        if triple not in seen:
            seen.add(triple)
            cleaned.append(row)


    # REMOVE DUPLICATE COMMENTS (Keep only 1 per unique context)

    unique_records = []
    seen_context = set()

    for row in cleaned:
        context = row.get("context", "")
        if context not in seen_context:
            seen_context.add(context)
            unique_records.append(row)

    cleaned = unique_records

    # EXPORT TO EXCEL

    print("Saving to Excel...")
    df = pd.DataFrame(cleaned)
    df.index += 1
    df = df.rename(columns={"key": "Key", "value": "Value", "context": "Comments"})
    df.to_excel(OUTPUT_EXCEL, index_label="#")

    print("Excel Generated:", OUTPUT_EXCEL)

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
    df.to_excel(writer, index_label="#", sheet_name="Output")

    sheet = writer.sheets["Output"]
    sheet.column_dimensions['A'].width = 4
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 35
    sheet.column_dimensions['D'].width = 100

    # Bold Header
    for cell in sheet["1:1"]:
        cell.font = cell.font.copy(bold=True)
