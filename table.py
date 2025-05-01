#Run with "streamlit run table.py" in the terminal after debug, start with "python -m" if using VS Code

import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import csv
import io
import difflib
import hashlib

def save_to_excel_heel(df, sheet_name='HEEL KITS', file_path='heel_data.xlsx'):
    if not os.path.exists(file_path):
        # Create a new file with the sheet
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Open existing and update/replace the sheet
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def save_to_excel_upper(df, sheet_name='HEEL KITS', file_path='upper_data.xlsx'):
    if not os.path.exists(file_path):
        # Create a new file with the sheet
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Open existing and update/replace the sheet
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)



st.set_page_config(page_title="SKU Kit Explorer", layout="wide")
st.title("👟 SKU Kit Explorer")

# --- Preload the Excel File
EXCEL_PATH = "MASTER Pricing & Skus.xlsx"

if not os.path.exists(EXCEL_PATH):
    st.error(f"❌ File '{EXCEL_PATH}' not found. Please make sure it's in the app directory.")
    st.stop()

heel_df = pd.read_excel(EXCEL_PATH, sheet_name="HEEL KITS")
st.session_state["shoe_df"] = pd.read_excel(EXCEL_PATH, sheet_name="SHOES")

st.subheader("👠 HEEL KITS")

# --- Ensure heel_df is in session state
if "heel_df" not in st.session_state:
    st.session_state["heel_df"] = heel_df.copy()



def file_hash(file):
    return hashlib.md5(file.getvalue()).hexdigest()

# Store uploaded hashes
if "uploaded_heel_hashes" not in st.session_state:
    st.session_state.uploaded_heel_hashes = set()

with st.expander("➕ Add New Heel SKU"):
    uploaded_csv = st.file_uploader("Upload CSV file", type=["csv"], key="csv_heel_upload")
    heel_columns = list(st.session_state["heel_df"].columns)
    required_heel_cols = ["PARENT SKU", "HEEL TYPE", "COLOR", "HEEL HEIGHT"]

    if uploaded_csv:
        current_hash = file_hash(uploaded_csv)

        if current_hash in st.session_state.uploaded_heel_hashes:
            st.info("ℹ️ This file has already been uploaded.")
        else:
            try:
                uploaded_df = pd.read_csv(uploaded_csv)

                missing_cols = [col for col in required_heel_cols if col not in uploaded_df.columns]
                if missing_cols:
                    st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
                else:
                    allowed_cols = [col for col in uploaded_df.columns if col in heel_columns]
                    clean_df = uploaded_df[allowed_cols].copy()
                    clean_df = clean_df.reindex(columns=heel_columns, fill_value="N/A")

                    st.session_state["heel_df"] = pd.concat([
                        st.session_state["heel_df"],
                        clean_df
                    ], ignore_index=True)

                    # Mark this file as processed
                    st.session_state.uploaded_heel_hashes.add(current_hash)

                    save_to_excel_heel(st.session_state["heel_df"])
                    st.success(f"✅ Imported {len(clean_df)} new heel SKUs. Columns added: {', '.join(allowed_cols)}")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ Error processing uploaded CSV: {e}")


    st.markdown("### 📁 Or Import Heel SKUs manually")
    new_heel_sku = st.text_input("Heel SKU")
    new_heel_type = st.text_input("Heel Type")
    new_heel_color = st.text_input("Heel Color")
    new_heel_height = st.text_input("Heel Height")

    if st.button("Add Heel"):
        if new_heel_sku:
            st.session_state["heel_df"] = pd.concat([
                st.session_state["heel_df"],
                pd.DataFrame([{
                    'PARENT SKU': new_heel_sku.strip(),
                    'HEEL TYPE': new_heel_type.strip(),
                    'COLOR': new_heel_color.strip(),
                    'HEEL HEIGHT': new_heel_height.strip()
                }])
            ], ignore_index=True)
            save_to_excel_heel(st.session_state["heel_df"])
            st.success(f"✅ Added new heel: {new_heel_sku}")
            st.rerun()
        else:
            st.warning("⚠️ Heel SKU is required.")

# --- Remove heel
with st.expander("❌ Remove Heel by SKU"):
    remove_heel_sku = st.text_input("Enter Heel SKU to Remove")
    if st.button("Remove Heel"):
        before = st.session_state["heel_df"].shape[0]
        st.session_state["heel_df"] = st.session_state["heel_df"][
            st.session_state["heel_df"]['PARENT SKU'] != remove_heel_sku.strip()
        ]
        after = st.session_state["heel_df"].shape[0]
        if before != after:
            save_to_excel_heel(st.session_state["heel_df"])
            st.success(f"🗑️ Removed heel: {remove_heel_sku}")
            st.rerun()
        else:
            st.warning("❗ Heel SKU not found.")

# --- Show current heel table
st.dataframe(st.session_state["heel_df"], use_container_width=True)



# UPPERS
st.subheader("👟 Shoe Uppers")
with st.expander("➕ Add New Upper SKU"):
   import hashlib

def file_hash(file):
    return hashlib.md5(file.getvalue()).hexdigest()

# Track uploaded shoe CSVs
if "uploaded_shoe_hashes" not in st.session_state:
    st.session_state.uploaded_shoe_hashes = set()

with st.expander("➕ Add New Upper SKU"):
    uploaded_csv = st.file_uploader("Upload CSV file", type=["csv"], key="csv_upper_upload")
    df_columns = list(st.session_state["shoe_df"].columns)
    required_cols = ["PARENT SKU", "SILHOUETTE", "UPPER NAME", "BASE COLOR", "MATERIAL CATEGORY"]

    if uploaded_csv:
        current_hash = file_hash(uploaded_csv)

        if current_hash in st.session_state.uploaded_shoe_hashes:
            st.info("ℹ️ This file has already been uploaded.")
        else:
            try:
                uploaded_df = pd.read_csv(uploaded_csv)

                # Check for required columns
                missing_cols = [col for col in required_cols if col not in uploaded_df.columns]
                if missing_cols:
                    st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
                else:
                    # Only keep columns that exist in shoe_df
                    allowed_cols = [col for col in uploaded_df.columns if col in df_columns]
                    clean_df = uploaded_df[allowed_cols].copy()

                    # Fill missing values with "N/A" for consistent format
                    clean_df = clean_df.reindex(columns=df_columns, fill_value="N/A")

                    # Merge into session state
                    st.session_state["shoe_df"] = pd.concat([
                        st.session_state["shoe_df"],
                        clean_df
                    ], ignore_index=True)

                    # Mark file as uploaded
                    st.session_state.uploaded_shoe_hashes.add(current_hash)

                    save_to_excel_upper(st.session_state["shoe_df"])
                    st.success(f"✅ Imported {len(clean_df)} new shoe SKUs. Columns added: {', '.join(allowed_cols)}")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ Error processing uploaded CSV: {e}")


    st.markdown("### 📁 Or Import Upper SKUs manually")
    new_upper_sku = st.text_input("Upper SKU")
    new_upper_silhouette = st.text_input("Silhouette")
    new_upper_name = st.text_input("Upper Name")
    new_upper_color = st.text_input("Base Color")
    new_upper_material = st.text_input("Material Category")
    
    if st.button("Add Upper"):
        if new_upper_sku:
            st.session_state["shoe_df"] = pd.concat([
                st.session_state["shoe_df"],
                pd.DataFrame([{
                    'PARENT SKU': new_upper_sku.strip(),
                    'SILHOUETTE': new_upper_silhouette.strip(),
                    'UPPER NAME': new_upper_name.strip(),
                    'BASE COLOR': new_upper_color.strip(),
                    'MATERIAL CATEGORY': new_upper_material.strip()
                }])
            ], ignore_index=True)
            st.success(f"✅ Added new upper: {new_upper_sku}")
        else:
            st.warning("⚠️ SKU is required.")

with st.expander("❌ Remove Upper by SKU"):
    remove_upper_sku = st.text_input("Enter Upper SKU to Remove")
    if st.button("Remove Upper"):
        before = st.session_state["shoe_df"].shape[0]
        st.session_state["shoe_df"] = st.session_state["shoe_df"][st.session_state["shoe_df"]['PARENT SKU'] != remove_upper_sku.strip()]
        after = st.session_state["shoe_df"].shape[0]
        if before != after:
            st.success(f"🗑️ Removed upper: {remove_upper_sku}")
        else:
            st.warning("❗ SKU not found.")

st.dataframe(st.session_state["shoe_df"], use_container_width=True)




# --- Sync to Upper Dictionary for Kit Matching
uppers = {}
for _, row in st.session_state["shoe_df"].iterrows():
    sku = row['PARENT SKU']
    if pd.notna(sku):
        uppers[sku.strip()] = {
            'SILHOUETTE': row['SILHOUETTE'].strip() if pd.notna(row['SILHOUETTE']) else '',
            'UPPER NAME': row['UPPER NAME'].strip() if pd.notna(row['UPPER NAME']) else '',
            'BASE COLOR': row['BASE COLOR'].strip() if pd.notna(row['BASE COLOR']) else '',
            'MATERIAL CATEGORY': str(row['MATERIAL CATEGORY']).strip() if pd.notna(row['MATERIAL CATEGORY']) else ''
        }


# --- Pre-process Heel and Upper data
heels = {}
for _, row in heel_df.iterrows():
    sku = row['PARENT SKU']
    if pd.notna(sku):
        heels[sku.strip()] = {
            'TYPE': row['HEEL TYPE'].strip() if pd.notna(row['HEEL TYPE']) else '',
            'COLOR': row['COLOR'].strip() if pd.notna(row['COLOR']) else '',
            'HEEL HEIGHT': str(row['HEEL HEIGHT']).strip() if pd.notna(row['HEEL HEIGHT']) else ''
        }

uppers = {}
for _, row in st.session_state["shoe_df"].iterrows():
    sku = row['PARENT SKU']
    if pd.notna(sku):
        uppers[sku.strip()] = {
            'SILHOUETTE': row['SILHOUETTE'].strip() if pd.notna(row['SILHOUETTE']) else '',
            'UPPER NAME': row['UPPER NAME'].strip() if pd.notna(row['UPPER NAME']) else '',
            'BASE COLOR': row['BASE COLOR'].strip() if pd.notna(row['BASE COLOR']) else '',
            'MATERIAL CATEGORY': str(row['MATERIAL CATEGORY']).strip() if pd.notna(row['MATERIAL CATEGORY']) else ''
        }

# --- File Upload
csv_file = st.file_uploader("📥 Upload Shopify Product CSV", type=["csv"])

# --- Match Kit SKU Logic
def infer_most_similar(candidate, options):
    # Return the best fuzzy match if similarity is high enough
    matches = difflib.get_close_matches(candidate, options, n=1, cutoff=0.6)
    return matches[0] if matches else None

def process_kit_sku(kit, uppers, heels):
    match = re.match(r'^(\d+)', kit)
    shoe_size = match.group(1) if match else None
    size = int(shoe_size)
    if size>11:
        size/=10
    remaining = kit[len(shoe_size):] if shoe_size else kit

    for upper_key in uppers:
        if remaining.startswith(upper_key):
            heel_candidate = remaining[len(upper_key):]
        else:
            upper_key=remaining[:-4]
            heel_candidate = remaining[len(upper_key):]
            for heel_key in heels:
                if heel_key.endswith(heel_candidate):
                    return kit, upper_key, heel_key, size
    return kit, None, None, size

    

if csv_file:
    # --- Step 1: Read & clean CSV
    df = pd.read_csv(csv_file)
    useful_cols = ['Handle', 'Title', 'Type', 'Option1 Name', 'Option1 Value', 'Variant SKU', 'Status']
    df = df[useful_cols]
    df[['Title', 'Type', 'Option1 Name', 'Status','Variant SKU', 'Option1 Value']] = (
        df.groupby('Handle')[['Title', 'Type', 'Option1 Name', 'Status','Variant SKU', 'Option1 Value']].ffill().bfill()
    )
    df = df.dropna(subset=['Variant SKU']).drop_duplicates()
    df['Type'] = df['Type'].replace('Custom Pairings', 'Custom Pairing')
    kit_skus = set(df[df['Type'].isin(['Shoes', 'Custom Pairing'])]['Variant SKU'])


    

    # --- Step 2: Match kits
    df_complete_list = []
    df_missing_list = []

    for kit in kit_skus:
        kit_val, upper_val, heel_val, size = process_kit_sku(kit, uppers, heels)
        row = {
            "KitSKU": kit_val,
            "ShoeSize": size,
            "UpperSKU": upper_val,
            "UpperDescription": uppers.get(upper_val) if upper_val else None,
            "HeelSKU": heel_val,
            "HeelDescription": heels.get(heel_val) if heel_val else None
        }
        (df_complete_list if upper_val and heel_val else df_missing_list).append(row)

    df_cheatsheet = pd.DataFrame(df_complete_list)
    df_missing = pd.DataFrame(df_missing_list)

    # --- Expand dict columns
    for col, prefix in [("UpperDescription", "Upper_"), ("HeelDescription", "Heel_")]:
        if col in df_cheatsheet.columns:
            expanded = df_cheatsheet[col].apply(lambda x: pd.Series(x) if isinstance(x, dict) else pd.Series())
            expanded = expanded.add_prefix(prefix)
            df_cheatsheet = df_cheatsheet.drop(col, axis=1).join(expanded)

    st.success("✅ Processing complete! Choose a view:")


# Function to convert DataFrame to Excel bytes
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Tabs to display and download results
tab1, tab2 = st.tabs(["🎯 Complete Kits", "⚠️ Missing Info"])

with tab1:
    if "df_cheatsheet" in locals():
        st.dataframe(df_cheatsheet, use_container_width=True)
        excel_data = convert_df_to_excel(df_cheatsheet)
        st.download_button(
            label="⬇️ Download Complete Kits (Excel)",
            data=excel_data,
            file_name="cheatsheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Upload and process product CSV to see complete kits.")

with tab2:
    if "df_missing" in locals():
        st.dataframe(df_missing, use_container_width=True)
        excel_missing = convert_df_to_excel(df_missing)
        st.download_button(
            label="⬇️ Download Missing Kits (Excel)",
            data=excel_missing,
            file_name="missing_cheatsheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Upload and process product CSV to see missing kits.")

