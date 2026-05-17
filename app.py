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
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt
        )
        
        response_text = response.text
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        
        if json_match:
            return json.loads(json_match.group(0))
        else:
            st.warning("Could not isolate JSON data from the AI response.")
            return {"vendor_name": "Unknown Vendor", "prices": {}}
            
    except Exception as e:
        st.error(f"Error communicating with Google AI: {e}")
        return {"vendor_name": "Unknown Vendor", "prices": {}}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Processing documents..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            items = []
            row_mapping = {}
            
            # Read items from Column C starting from row 8 (or lower if headers shift)
            # Since Vendor Names are on Row 8, let's start reading item lines from Row 9 downwards
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value # Column 3 is C
                if item_desc and str(item_desc).strip():
                    clean_desc = str(item_desc).strip()
                    # Skip header noise if any accidental rows are parsed
                    if "ITEM DESCRIPTION" in clean_desc.upper():
                        continue
                    items.append(clean_desc)
                    row_mapping[clean_desc] = row
            
            if not items:
                st.error("No item descriptions found in Column C (checked from row 9 downwards).")
            else:
                st.info(f"Found {len(items)} items to look up prices for.")
                
                # Excel column tracks:
                # Vendor 1: Unit Price = E (5), Name = F (6)
                # Vendor 2: Unit Price = H (8), Name = I (9)
                # Vendor 3: Unit Price = K (11), Name = L (12)
                # Vendor 4: Unit Price = N (14), Name = O (15)
                vendor_unit_cols = [5, 8, 11, 14]
                vendor_name_cols = [6, 9, 12, 15]
                
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"🔄 Processing PFI: **{pfi.name}**...")
                    
                    pfi_text = extract_text_from_pdf(pfi)
                    extracted_data = parse_pfi_with_ai(pfi_text, items, api_key)
                    
                    vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
                    extracted_prices = extracted_data.get("prices", {})
                    
                    unit_col = vendor_unit_cols[index]
                    name_col = vendor_name_cols[index]
                    
                    # 1. Insert Vendor Name in Row 8 of the corresponding column (F, I, L, or O)
                    ws.cell(row=8, column=name_col).value = vendor_name
                    
                    # 2. Insert Unit Prices into column E, H, K, or N
                    match_count = 0
                    for item, price in extracted_prices.items():
                        if item in row_mapping:
                            target_row = row_mapping[item]
                            try:
                                ws.cell(row=target_row, column=unit_col).value = float(price)
                                match_count += 1
                            except ValueError:
                                pass
                    
                    st.success(f"✅ Finished **{vendor_name}** ({pfi.name}). Matched {match_count}/{len(items)} prices.")
                
                output_filename = "Populated_Comparative_Analysis.xlsx"
                wb.save(output_filename)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Completed Excel Sheet",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                os.remove(output_filename)
