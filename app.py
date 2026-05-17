import streamlit as st
import pypdf
import openpyxl
from google import genai
import json
import os

st.set_page_config(page_title="Procurement Automation Tool", layout="wide")
st.title("📊 Automated Vendor PFI Comparative Analysis")
st.subheader("Upload your template and vendor PFIs to auto-populate prices.")

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
    pfi_files = st.file_uploader("2. Upload Vendor PFIs (PDFs)", type=["pdf"], accept_multiple_files=True)

def extract_text_from_pdf(pdf_file):
    """Extracts raw text from an uploaded PDF file."""
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def parse_pfi_with_ai(pfi_text, items_list, api_key_str):
    """Uses Gemini to find the unit prices for our target items from the raw PFI text."""
    client = genai.Client(api_key=api_key_str)
    
    prompt = f"""
    You are an expert procurement data extraction tool.
    I am going to give you the raw text extracted from a Vendor's Proforma Invoice (PFI).
    I want you to find the UNIT PRICE for the following target items.
    
    Target Items to look for:
    {json.dumps(items_list)}
    
    Instructions:
    1. Look closely at the PFI text. Match the items even if the vendor's description is slightly different or abbreviated.
    2. Extract the Unit Price for each matched item.
    3. Return the result STRICTLY as a JSON dictionary where the keys are the EXACT item descriptions from the target list, and values are numbers (floats) representing the unit price. If an item is not found in the PFI, omit it from the dictionary.
    
    Raw PFI Text:
    \"\"\"{pfi_text}\"\"\"
    """
    
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=prompt,
        config={'response_mime_type': 'application/json'}
    )
    
    try:
        return json.loads(response.text)
    except Exception as e:
        st.error(f"Failed to parse AI response: {e}")
        return {}

if st.button("🚀 Process Invoices & Match Prices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Processing documents..."):
            # 1. Load the Excel Template
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active # Assumes first sheet
            
            # 2. Read the items from Column C (Item Description), starting from row 8
            items = []
            row_mapping = {} # To remember which item belongs to which Excel row
            
            for row in range(8, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value # Column 3 is C
                if item_desc and str(item_desc).strip():
                    clean_desc = str(item_desc).strip()
                    items.append(clean_desc)
                    row_mapping[clean_desc] = row
            
            if not items:
                st.error("No item descriptions found in Column C (starting row 8). Check your template format.")
            else:
                st.info(f"Found {len(items)} items to look up prices for.")
                
                # 3. Define the column index mapping for vendors
                # Vendor 1 Unit Price = Col E (5)
                # Vendor 2 Unit Price = Col H (8)
                # Vendor 3 Unit Price = Col K (11)
                # Vendor 4 Unit Price = Col N (14)
                vendor_unit_cols = [5, 8, 11, 14]
                
                # Loop through each uploaded PFI file (up to 4)
                for index, pfi in enumerate(pfi_files[:4]):
                    st.write(f"🔄 Processing PFI: **{pfi.name}** as Vendor {index + 1}...")
                    
                    # Extract text and send to Gemini
                    pfi_text = extract_text_from_pdf(pfi)
                    extracted_prices = parse_pfi_with_ai(pfi_text, items, api_key)
                    
                    # Target column for this vendor's unit price
                    target_col = vendor_unit_cols[index]
                    
                    # Write prices into the spreadsheet
                    match_count = 0
                    for item, price in extracted_prices.items():
                        if item in row_mapping:
                            target_row = row_mapping[item]
                            ws.cell(row=target_row, column=target_col).value = float(price)
                            match_count += 1
                    
                    st.success(f"✅ Finished {pfi.name}. Matched {match_count}/{len(items)} prices.")
                
                # 4. Save file back to binary memory so the user can download it
                output_filename = "Populated_Comparative_Analysis.xlsx"
                wb.save(output_filename)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Completed Excel Sheet",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Clean up local file copy
                os.remove(output_filename)
