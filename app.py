import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Procurement Automation Tool", layout="wide")
st.title("📊 Automated Vendor PFI Comparative Analysis (Diagnostic Engine)")

# Sidebar for API Key configuration
with st.sidebar:
    st.header("Setup")
    api_key = st.text_input("Enter Gemini API Key:", type="password")

template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
pfi_files = st.file_uploader("2. Upload Vendor PFIs (PDFs)", type=["pdf"], accept_multiple_files=True)

def convert_pdf_to_images(pdf_file):
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        return convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"Failed to process PDF pages into images: {e}")
        return []

def parse_pfi_images_with_ai(pfi_images, row_item_dict, api_key_str):
    try:
        client = genai.Client(api_key=api_key_str)
        
        prompt_text = f"""You are a strict data extraction tool. Look at the attached invoice image(s).
        
        Our Master Spreadsheet Items (Format is "ROW_NUMBER": "ITEM DESCRIPTION"):
        {json.dumps(row_item_dict, indent=2)}
        
        Task:
        1. Extract the Vendor Name from the letterhead.
        2. Match the invoice items to our row list. Extract ONLY the raw numerical UNIT PRICE. Clean out currency symbols and commas.
        
        Return your answer strictly as a clean JSON object structure (no markdown blocks):
        {{
            "vendor_name": "Official Vendor Name",
            "row_prices": {{
                "9": 1500.00
            }}
        }}"""
        
        contents = [prompt_text]
        for img in pfi_images:
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        
        return json.loads(response.text.strip())
    except Exception as e:
        st.error(f"AI Extraction Failed: {e}")
        return {"vendor_name": "Error", "row_prices": {}}

if st.button("🚀 Process Invoices") and template_file and pfi_files:
    if not api_key:
        st.warning("Please enter your Gemini API Key.")
    else:
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        
        row_item_dict = {}
        for row in range(9, ws.max_row + 1):
            item_desc = ws.cell(row=row, column=3).value
            if item_desc:
                clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                    continue
                row_item_dict[str(row)] = clean_desc
        
        vendor_start_cols = [5, 8, 11, 14] # E, H, K, N
        
        for index, pfi in enumerate(pfi_files[:4]):
            st.write(f"### Diagnostic Processing for: **{pfi.name}**")
            pfi_images = convert_pdf_to_images(pfi)
            
            if not pfi_images:
                continue
                
            extracted_data = parse_pfi_images_with_ai(pfi_images, row_item_dict, api_key)
            
            # --- CRITICAL LIVE DIAGNOSTIC BOX ---
            st.info("📦 **Raw JSON Extracted by Gemini:**")
            st.json(extracted_data)
            
            vendor_name = extracted_data.get("vendor_name", f"Vendor {index + 1}")
            row_prices = extracted_data.get("row_prices", {})
            
            target_col = vendor_start_cols[index]
            ws.cell(row=8, column=target_col).value = vendor_name
            
            st.write("🔧 **Python Data Insertion Logs:**")
            
            match_count = 0
            for row_str, price in row_prices.items():
                try:
                    target_row = int(row_str)
                    clean_price = float(str(price).replace(",", "").strip())
                    
                    # Direct assignment
                    ws.cell(row=target_row, column=target_col).value = clean_price
                    st.write(f"➡️ Row `{target_row}` Column `{target_col}` successfully populated with `{clean_price}`")
                    match_count += 1
                except Exception as cell_error:
                    # If any assignment breaks, don't hide it! Tell us why!
                    st.error(f"❌ Failed to write row {row_str} with value '{price}': {cell_error}")
            
            st.success(f"Calculated {match_count} updates for {vendor_name}")
        
        output_filename = "Populated_Comparative_Analysis.xlsx"
        wb.save(output_filename)
        st.download_button(label="📥 Download Output Sheet", data=open(output_filename, "rb"), file_name=output_filename)
        os.remove(output_filename)
