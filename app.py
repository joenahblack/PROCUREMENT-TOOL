import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Procurement Tool - Single Vendor Engine", layout="wide")
st.title("📊 Single Vendor Quotation Matcher")
st.subheader("Upload your template and one vendor quotation to map material prices cleanly.")

# Sidebar for Setup
with st.sidebar:
    st.header("Authentication")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get a free API Key here](https://aistudio.google.com/)")

# File Uploaders
template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
pfi_file = st.file_uploader("2. Upload Single Vendor Quotation (PDF)", type=["pdf"])

def convert_pdf_to_images(pdf_file):
    """Converts the PDF into images so the AI can physically see scanned text and tables."""
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        return convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"Failed to convert PDF to images. Is Poppler configured? Error: {e}")
        return []

if st.button("🚀 Match & Populate Excel") and template_file and pfi_file:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Processing document and running visual alignment..."):
            # 1. Open the Excel workbook
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            # 2. Build our target mapping map directly from Column C (Row 9 downwards)
            row_item_dict = {}
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value # Column C is 3
                if item_desc:
                    clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                    # Skip header labels or structural noise
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    row_item_dict[str(row)] = clean_desc
            
            st.info(f"Loaded {len(row_item_dict)} material lines from your Excel sheet.")
            
            # 3. Convert PDF pages to images
            images = convert_pdf_to_images(pfi_file)
            
            if not images:
                st.error("Could not extract pages from the PDF quotation.")
            else:
                # 4. Craft the prompt forcing Gemini to map using spreadsheet rows directly
                prompt_text = f"""You are an expert procurement data extraction engine. Analyze the attached quotation image(s).
                
                Tasks:
                1. Identify the official VENDOR NAME from the header/letterhead.
                2. Match the items on this invoice to our target spreadsheet items list. 
                   - Look past severe industry abbreviations (e.g., 'BLT' -> 'Bolt', 'SS' -> 'Stainless Steel'). Match them flexibly using your procurement domain knowledge.
                3. Extract ONLY the raw numeric UNIT PRICE. Ignore quantities or line totals.
                
                Our Target Spreadsheet Items (Format is "ROW_NUMBER": "EXACT MATERIAL DESCRIPTION"):
                {json.dumps(row_item_dict, indent=2)}
                
                Return your response strictly as a clean JSON object structure matching this exact format. Do not use markdown syntax block wraps:
                {{
                    "vendor_name": "Official Vendor Name",
                    "row_prices": {{
                        "9": 1500.00
                    }}
                }}"""
                
                # Assemble text instructions and visual pages
                contents = [prompt_text]
                for img in images:
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG')
                    contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
                
                try:
                    # Initialize the official Google GenAI Client
                    client = genai.Client(api_key=api_key)
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=contents,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    
                    response_text = response.text.strip()
                    extracted_data = json.loads(response_text)
                    
                    # --- LIVE DIAGNOSTIC STREAM ON SCREEN ---
                    st.success("🤖 **Raw Live AI Extraction Data Matrix:**")
                    st.json(extracted_data)
                    
                    vendor_name = extracted_data.get("vendor_name", "Unknown Vendor")
                    row_prices = extracted_data.get("row_prices", {})
                    
                    # 5. Write Vendor Name to Column E, Row 8 (First vendor slot)
                    ws.cell(row=8, column=5).value = vendor_name
                    
                    # 6. Inject prices down Column E using the exact row keys provided by the AI
                    match_count = 0
                    st.write("### ⚙️ Excel Writing Stream Logs:")
                    for row_str, price in row_prices.items():
                        try:
                            target_row = int(row_str)
                            # Strip formatting remnants safely
                            clean_price = float(str(price).replace(",", "").strip())
                            
                            ws.cell(row=target_row, column=5).value = clean_price
                            st.write(f"➡️ Placed price `{clean_price}` onto Row `{target_row}` (Material: *{row_item_dict.get(row_str)}*)")
                            match_count += 1
                        except (ValueError, TypeError):
                            pass
                            
                    st.success(f"🎉 Complete! Mapped {match_count} material items for **{vendor_name}** into Column E.")
                    
                    # 7. Package updated file for instant browser download
                    output_filename = "Single_Vendor_Populated.xlsx"
                    wb.save(output_filename)
                    
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label="📥 Download Processed Excel Sheet",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    os.remove(output_filename)
                    
                except Exception as e:
                    st.error(f"Error communicating with Gemini Vision API: {e}")
