import streamlit as st
import openpyxl
from google import genai
from google.genai import types
from pdf2image import convert_from_bytes
import io
import json
import os
import re

st.set_page_config(page_title="Procurement Tool - Alignment Engine", layout="wide")
st.title("📊 Part-Quote Alignment Matcher")
st.subheader("Maps available vendor prices to correct material cells, automatically skipping unquoted items.")

# Sidebar for Setup
with st.sidebar:
    st.header("Authentication")
    api_key = st.text_input("Enter Gemini API Key:", type="password")
    st.markdown("[Get a free API Key here](https://aistudio.google.com/)")

# File Uploaders
template_file = st.file_uploader("1. Upload Excel Template (.xlsx)", type=["xlsx"])
pfi_file = st.file_uploader("2. Upload Single Vendor Quotation (PDF)", type=["pdf"])

def convert_pdf_to_images(pdf_file):
    try:
        pdf_bytes = pdf_file.read()
        pdf_file.seek(0)
        return convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"Failed to convert PDF to images: {e}")
        return []

if st.button("🚀 Match & Populate Excel") and template_file and pfi_file:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        with st.spinner("Aligning partial quotation matrices..."):
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            # Read all master materials from Column C (Row 9 down)
            row_item_dict = {}
            for row in range(9, ws.max_row + 1):
                item_desc = ws.cell(row=row, column=3).value
                if item_desc:
                    clean_desc = re.sub(r'\s+', ' ', str(item_desc)).strip()
                    if "ITEM DESCRIPTION" in clean_desc.upper() or not clean_desc:
                        continue
                    row_item_dict[str(row)] = clean_desc
            
            st.info(f"Loaded {len(row_item_dict)} total material items requested from Excel.")
            
            images = convert_pdf_to_images(pfi_file)
            
            if not images:
                st.error("Could not process PDF pages.")
            else:
                prompt_text = f"""You are a precise procurement data entry operator. Analyze the attached vendor quotation image(s).
                
                CRITICAL INSTRUCTIONS FOR PARTIAL QUOTES:
                1. Identify the official VENDOR NAME from the document header.
                2. Cross-reference the items on the invoice with our Master Spreadsheet Items list below.
                3. The vendor may NOT have quoted for all items requested. For every item the vendor DID quote, find its matching item in our Master list and extract its numeric UNIT PRICE.
                4. If an item on our Master list is missing from the vendor invoice, DO NOT include its row number in the output dictionary.
                5. Ensure you extract the UNIT PRICE, not the quantity or line total.
                
                Our Master Spreadsheet Items (Format is \"ROW_NUMBER\": \"EXACT MATERIAL DESCRIPTION\"):
                {json.dumps(row_item_dict, indent=2)}
                
                Return your response strictly as a clean JSON object structure (no markdown wrappers):
                {{
                    "vendor_name": "Official Vendor Name",
                    "row_prices": {{
                        "SPREADSHEET_ROW_NUMBER": 1500.00
                    }}
                }}"""
                
                contents = [prompt_text]
                for img in images:
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG')
                    contents.append(types.Part.from_bytes(data=img_byte_arr.getvalue(), mime_type="image/jpeg"))
                
                try:
                    client = genai.Client(api_key=api_key)
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=contents,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    
                    extracted_data = json.loads(response.text.strip())
                    
                    # Live status screen print
                    st.success("🤖 **Live AI Extraction Matrix Map:**")
                    st.json(extracted_data)
                    
                    vendor_name = extracted_data.get("vendor_name", "Unknown Vendor")
                    row_prices = extracted_data.get("row_prices", {})
                    
                    # Place Vendor Name in Column E, Row 8
                    ws.cell(row=8, column=5).value = vendor_name
                    
                    # Inject prices precisely down Column E matching row designations
                    match_count = 0
                    st.write("### ⚙️ Cell Allocation Log:")
                    for row_str, price in row_prices.items():
                        try:
                            target_row = int(row_str)
                            clean_price = float(str(price).replace(",", "").strip())
                            
                            # Input value into cell coordinate
                            ws.cell(row=target_row, column=5).value = clean_price
                            st.write(f"➡️ Injected `{clean_price}` into Cell `E{target_row}` (Material Match: *{row_item_dict.get(row_str)}*)")
                            match_count += 1
                        except (ValueError, TypeError):
                            pass
                            
                    st.success(f"🎉 Aligned and mapped {match_count} prices for **{vendor_name}** into Column E. Unquoted rows were safely skipped.")
                    
                    output_filename = "Aligned_Vendor_Analysis.xlsx"
                    wb.save(output_filename)
                    
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label="📥 Download Aligned Spreadsheet",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    os.remove(output_filename)
                    
                except Exception as e:
                    st.error(f"API Communication Error: {e}")
