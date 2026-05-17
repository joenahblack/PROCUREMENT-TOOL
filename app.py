import streamlit as st
import pypdf
import openpyxl
from google import genai
import json
import os
import re

st.set_page_config(page_title="Procurement Automation Tool", layout="wide")
st.title("📊 Automated Vendor PFI Comparative Analysis")
st.subheader("Upload your template and vendor PFIs to auto-populate names and prices.")

# Sidebar for API Key configuration
with st.sidebar:
    st.header("Setup")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get a free API Key here](https://aistudio.google.com/)")

# File Uploaders
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
with col2:
    pfi_files = st.file_uploader("2. Upload ALL Vendor PFIs together (PDFs)", type=["pdf"], accept_multiple_files=True)

def extract_text_from_pdf(pdf_file):
    """Extracts raw text from an uploaded PDF file."""
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def parse_pfi_with_ai(pfi_text, items_list, api_key_str):
    """Uses Gemini to extract Vendor Name and Unit Prices from the raw PFI text."""
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt = f"""
        You are an expert procurement data extraction tool.
        Look closely at the raw text extracted from a Vendor's Proforma Invoice (PFI).
        
        Tasks:
        1. Identify the official VENDOR NAME (usually found at the top of the PFI).
        2. Find the UNIT PRICE for the following target items. Match them even if descriptions are slightly different or abbreviated.
        
        Target Items to look for:
        {json.dumps(items_list)}
        
        Instructions:
        Return your answer STRICTLY as a raw JSON dictionary with exactly two keys: "vendor_name" and "prices". 
        The "prices" key must be a dictionary where keys are the EXACT item descriptions from the target list, and values are numbers (floats) representing the unit price. If an item is not found, omit it.
        
        Example Output format:
        {{
            "vendor_name": "ABC Supplies Ltd",
            "prices": {{
                "ITEM DESCRIPTION 1": 1500.50,
                "ITEM DESCRIPTION 2": 450.00
            }}
        }}

        Raw PFI Text:
        \"\"\"{pfi_text}\"\"\"
        """
        
        # Updated to utilize gemini-2.5-flash
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt
        )
        
        response_text = response.text
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        
        if json_match:
            return json.loads(json_match.group(0))
