# smart-document-extraction
AI-Powered pipeline that converts complex unstructured PDFs into structured Excel data with original context retention using Gemini LLM.
=======
# AI-Powered Document Structuring & Data Extraction

### Automated PDF to Structured Excel Converter

This project converts **unstructured PDF content into a clean, structured Excel sheet** using an **AI-backed extraction pipeline**.  
It was developed for the assignment titled:  
> **AI-Powered Document Structuring & Data Extraction Task**

---

## Project Objective

Build a system that:
- Extracts **100% content** from an unorganized document (PDF)
- Detects **logical key:value relationships**
- Preserves **original wording**
- Generates a structured Excel file similar to the provided `Expected Output.xlsx`

---

## Tech Used

| Technology | Purpose |
|------------|---------|
| Python | Core development |
| pdfplumber | Extract text from PDF |
| Google Gemini AI | Key:Value inference + context extraction |
| Pandas | Data cleaning & export |
| OpenPyXL | Excel formatting |
| Streamlit | Live demo web app |

---

## Features

✔ Extracts key:value pairs automatically  
✔ Preserves **original wording** (no paraphrasing)  
✔ Adds **Context column** from original PDF text  
✔ Removes duplicate & noisy data  
✔ Applies professional **Excel formatting**  
✔ Contains **live web demo**  

---

## Output Sample

| # | Key | Value | Comments |
|---|-----|-------|----------|
| 1 | Name | Vijay Kumar | Vijay Kumar was born on... |
| 2 | Place of Birth | Jaipur | Born and raised in... |
| … | … | … | … |

Full output generated: `Output.xlsx`

---

## How it Works

### Run the main automation script
#### Add your API key

follow the commands:
pip3 install -r requirements.txt
python3 main.py

#### You can also run a browser-based demo:
streamlit run streamlit_demo.py

Upload any PDF and download structured Excel instantly.

```bash
export GEMINI_API_KEY="your_key_here"
```

### Author

Ronak Vadher

Building practical AI tools for workflow automation & data transformation

3nd Year — Bsc Data Science
