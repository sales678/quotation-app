import os
import pandas as pd
import streamlit as st
   def get_salutation(name):
    name_upper = name.upper()
    
    # கம்பெனி பெயர்கள் (Company / Business)
    company_keywords = ["TRADERS", "MOTORS", "ENTERPRISES", "LIMITED", "LTD", "AGENCY", "WORKS", "STORES", "COMPANY", "CO"]
    if any(keyword in name_upper for keyword in company_keywords):
        return "M/S."
    
    # பெண்கள் பெயர்கள் (Female)
    female_keywords = ["MRS", "MISS", "KUMARI", "DEVI", "AMMAL", "MARY", "LAKSHMI", "ANITHA", "PRIYA", "KAVITHA"]
    if any(keyword in name_upper for keyword in female_keywords):
        return "MRS."
    
    # ஆண்கள் (Default Male)
    return "MR."
    
from docxtpl import DocxTemplate

st.set_page_config(page_title="Auto Vehicle Quotation Generator", layout="wide")
st.title("🚗 Auto Vehicle Quotation Generator")

# --- 1. EXCEL DATA LOAD ---
@st.cache_data
def load_excel_data():
    file_path = "QA.xlsx"
    if os.path.exists(file_path):
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_name = "PRICE" if "PRICE" in excel_file.sheet_names else excel_file.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = df.columns.astype(str).str.strip()
            return df
        except Exception as e:
            st.error(f"Excel படிக்க முடியவில்லை: {e}")
            return None
    return None

df = load_excel_data()

if df is None or df.empty:
    st.error("❌ 'QA.xlsx' ஃபைல் சரியாக லோட் ஆகவில்லை! ஃபைலை சரிபார்க்கவும்.")
    st.stop()

# --- 2. FSC DETAILS DATA ---
fsc_details = {
    "MURUGAN": "VQ126A0000000196",
    "SOWMIYAN": "VQ126A0000000155",
    "KANNAN": "VQ126A0000000156",
    "ARUN": "VQ126A0000000157",
    "MANIKANDAN": "VQ126A0000000158"
}

# --- 3. BANK LIST DATA ---
bank_options = [
    "HDFC BANK LTD",
    "ICICI BANK LTD",
    "AXIS BANK LTD",
    "STATE BANK OF INDIA",
    "SUNDARAM FINANCE LTD",
    "MAHINDRA FINANCE",
    "CHOLAMANDALAM FINANCE",
    "CANARA BANK",
    "INDIAN OVERSEAS BANK",
    "Other (Type Manually)"
]

# --- 4. VARIANT / MODEL COLUMN FINDER ---
# எக்செல்-ல Variant அல்லது Model Column-ஐ தானாக கண்டுபிடிக்கிறது
variant_column = None
for col in df.columns:
    col_str = str(col).upper()
    if any(k in col_str for k in ["VARIANT", "VEHICLE", "MODEL", "NAME", "DESCRIPTION"]):
        variant_column = col
        break

# ஒருவேளை கண்டுபிடிக்க முடியாவிட்டால் 1-வது காலமை எடுத்துக்கொள்ளும்
if not variant_column:
    variant_column = df.columns[0]

# --- 5. INPUT FIELDS ---
col1, col2 = st.columns(2)

# Document Upload Option for Auto-Fill
        uploaded_doc = st.file_uploader("📄 Upload Image/PDF (Aadhar/RC/Card)", type=["png", "jpg", "jpeg", "pdf"])
        
        default_name = "SK TRADERS"
        default_address = "100FT RING ROAD, HOSUR"
        
        # Name and Address Inputs
        cust_input = st.text_input("Customer Name", default_name)
        
        if cust_input:
            salutation = get_salutation(cust_input)
            customer_name = f"{salutation} {cust_input.upper()}"
        else:
            customer_name = ""
            
        customer_address = st.text_area("Customer Address", default_address)
    
    # FSC Name Selection
    fsc_name = st.selectbox("Select FSC Name", list(fsc_details.keys()))
    fsc_code = fsc_details[fsc_name]
    st.info(f"Selected FSC Code: **{fsc_code}**")

with col2:
    # 🚗 ஒரே Dropdown - இதில் வண்டி மாடல் / வேரியண்ட் லிஸ்ட் முழுமையாக வரும்!
    vehicle_variants = df[variant_column].dropna().unique()
    selected_variant = st.selectbox(" Select Vehicle Model / Variant", vehicle_variants)
    
    # 🏦 Bank Selection Box
    selected_bank_option = st.selectbox(" Select Hypothecation Bank", bank_options)
    
    if selected_bank_option == "Other (Type Manually)":
        bank_name = st.text_input("Enter Bank Name Manually", "HDFC BANK LTD")
    else:
        bank_name = selected_bank_option

# Selected Row Data
filtered_df = df[df[variant_column] == selected_variant]
model_row = filtered_df.iloc[0] if not filtered_df.empty else pd.Series()

st.divider()

# --- 6. PARTICULARS SELECT & AUTO-FILL PRICES ---
st.subheader("📋 Select Particulars & Auto-Filled Prices")

all_particulars = [
    "Ex Showroom", "Insurance", "KA LIFE TAX", "Life Tax", "Temp", 
    "RSA", "Yellow Paint & Handling", "Fast Tag", "Accessories", "TCS 1 %"
]

selected_particulars = st.multiselect(
    "Quotation-ல் வர வேண்டிய Particulars-ஐ தேர்வு செய்யவும்:",
    options=all_particulars,
    default=["Ex Showroom", "Insurance", "KA LIFE TAX", "TCS 1 %"]
)

items_data = []
total_cost = 0.0

col_p1, col_p2 = st.columns(2)

with col_p1:
    for item in selected_particulars:
        excel_val = 0.0
        
        if not model_row.empty:
            for col in df.columns:
                c_col = str(col).lower().replace(" ", "").replace("_", "")
                c_item = item.lower().replace(" ", "").replace("_", "")
                
                if c_item in c_col or c_col in c_item:
                    try:
                        val = model_row[col]
                        if pd.notna(val):
                            excel_val = float(val)
                    except:
                        excel_val = 0.0
                    break
        
        # Auto-Fill Price
        price = st.number_input(f"Price for {item} (₹)", value=excel_val, step=500.0, key=f"{selected_variant}_{item}")
        
        items_data.append({
            "particulars": item,
            "price": f"{int(price)}" if price.is_integer() else f"{price:,.2f}"
        })
        total_cost += price

formatted_total = f"{int(total_cost)}" if total_cost.is_integer() else f"{total_cost:,.2f}"
st.markdown(f"### 💰 **Total Cost: ₹ {formatted_total}**")

st.divider()

# --- 7. GENERATE WORD & PDF ---
if st.button("🚀 Generate Word & PDF Quotation"):
    template_path = "Quotation_Template.docx"

    if not os.path.exists(template_path):
        st.error("❌ 'Quotation_Template.docx' ஃபைல் ஃபால்டரில் இல்லை!")
    else:
        doc = DocxTemplate(template_path)

        context = {
            "date": pd.Timestamp.now().strftime("%d-%m-%Y"),
            "customer_name": customer_name,
            "customer_address": customer_address,
            "fsc_name": fsc_name,
            "fsc_code": fsc_code,
            "vehicle_variant": selected_variant,
            "bank_name": bank_name,
            "items": items_data,
            "total_cost": formatted_total
        }

        word_filename = f"Quotation_{customer_name}.docx"
        pdf_filename = f"Quotation_{customer_name}.pdf"

        doc.render(context)
        doc.save(word_filename)

        # PDF Conversion
        pdf_ready = False
        
        # docx2pdf conversion
        try:
            from docx2pdf import convert
            convert(word_filename, pdf_filename)
            pdf_ready = True
        except:
            try:
                import comtypes.client
                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = False
                doc_path = os.path.abspath(word_filename)
                pdf_path = os.path.abspath(pdf_filename)

                doc_pdf = word.Documents.Open(doc_path)
                doc_pdf.SaveAs(pdf_path, FileFormat=17) # 17 = PDF
                doc_pdf.Close()
                word.Quit()
                pdf_ready = True
            except Exception as e:
                st.warning(f"⚠️ PDF உருவாக்க முடியவில்லை: {e}")

        st.success("🎉 Quotation வெற்றிகரமாக உருவாக்கப்பட்டது!")

        c1, c2 = st.columns(2)
        
        with c1:
            with open(word_filename, "rb") as f_word:
                st.download_button(
                    label="📥 Download WORD File",
                    data=f_word,
                    file_name=word_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        with c2:
            if pdf_ready and os.path.exists(pdf_filename):
                with open(pdf_filename, "rb") as f_pdf:
                    st.download_button(
                        label="📄 Download PDF File",
                        data=f_pdf,
                        file_name=pdf_filename,
                        mime="application/pdf"
                    )
