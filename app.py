import os
import re
import gc
import io
import time
import zipfile
import shutil
import tempfile
import requests
import rarfile
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Manba Image Tool",
    page_icon="📥",
    layout="wide"
)

# --- GLOBAL SETTINGS & UTILS ---
INDIAN_TZ = timezone(timedelta(hours=5, minutes=30))
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

def sanitize_filename(name):
    """Removes illegal characters for Windows/Linux directory trees."""
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

def reset_metrics():
    st.session_state.processing_complete = False
    st.session_state.zip_path = ""
    st.session_state.success_count = 0
    st.session_state.failure_count = 0
    st.session_state.total_extracted_items = 0
    st.session_state.final_df_summary = None
    st.session_state.uploader_id = st.session_state.get('uploader_id', 0) + 1

# --- STATE CONFIGURATION SYNC ---
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False
if "zip_path" not in st.session_state:
    st.session_state.zip_path = ""
if "success_count" not in st.session_state:
    st.session_state.success_count = 0
if "failure_count" not in st.session_state:
    st.session_state.failure_count = 0
if "total_extracted_items" not in st.session_state:
    st.session_state.total_extracted_items = 0
if "final_df_summary" not in st.session_state:
    st.session_state.final_df_summary = None
if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0

# ====================================================
# DYNAMIC PREMIUM CSS STYLING
# ====================================================
st.markdown("""
    <style>
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { gap: 0.5rem !important; }
    section[data-testid="stSidebar"] { overflow-y: hidden !important; overflow-x: hidden !important; }
    
    .main-header { font-size: 34px; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .sub-header { font-size: 16px; color: #4B5563; margin-bottom: 25px; }
    .metric-box { background-color: #F3F4F6; padding: 18px; border-radius: 10px; border-left: 6px solid #2563EB; box-shadow: 0 2px 4px rgba(0,0,0,0.04); }
    .footer-text { text-align: center; font-size: 14px; color: #9CA3AF; margin-top: 50px; padding-top: 20px; border-top: 1px solid #E5E7EB; }
    .note-box { background-color: #FEF3C7; color: #92400E; padding: 12px; border-radius: 6px; border-left: 4px solid #F59E0B; margin-bottom: 15px; font-size: 14px; font-weight: 500; }
    
    .premium-card { background: #FFFFFF; padding: 24px; border-radius: 14px; border-bottom: 5px solid #CBD5E1; box-shadow: 0 4px 20px rgba(0,0,0,0.03); text-align: center; transition: transform 0.2s; }
    .premium-card:hover { transform: translateY(-2px); }
    .card-title { font-size: 12px; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
    .card-value { font-size: 38px; font-weight: 800; margin-top: 8px; }

    div[data-testid="stFileUploaderDropzoneInstructions"] small,
    div[data-testid="stFileUploaderDropzoneInstructions"] div,
    div[data-testid="stWidgetFormInstructions"],
    .stWidgetFormInstructions {
        display: none !important;
    }
    
    div[data-testid="stFileUploaderDropzoneInstructions"]::after {
        content: "⚠️ only xlsx file supported ";
        font-size: 13px;
        color: #DC2626;
        font-weight: 600;
        display: block;
        margin-top: 12px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR CONTROL PANEL ---
with st.sidebar:
    st.markdown("### 🎛️ Control Panel Matrix")
    st.info("📌 **Active Engine:** Bharath Cloud Extraction")
    st.write("---")
    st.markdown("##### 🔧 Workflow System Configurations")
    
    max_workers = st.slider("Concurrent Downloads", min_value=1, max_value=15, value=5)
    
    sample_data = {
        "Folder Name": ["CG19BR7427", "TN3006TW0162293"],
        "Link": [
            "https://manbafinance-bcsb3.bharathcloud.com:443/index.php/s/71HLqay3Nh0XJQ0/download",
            "https://manbafinance-bcsb3.bharathcloud.com/example/download"
        ]
    }
    sample_df = pd.DataFrame(sample_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False, sheet_name='Sheet1')
    st.download_button(
        label="📥 Download Excel Template",
        data=buffer.getvalue(),
        file_name="Download_Manba_Excel_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# --- HEADER & CONTEXT PANEL ---
st.markdown("## 📥 Manba Image Tool (Bharath Cloud RAR Extractor)")
st.markdown('<div class="note-box">📌 <strong>This Tool Works Only For Bharath Cloud RAR Links.</strong></div>', unsafe_allow_html=True)
st.markdown("🔗 **Sample Link Format:** `https://manbafinance-bcsb3.bharathcloud.com:443/index.php/s/...`")

st.write("")
layout_left, layout_right = st.columns([2, 1])

with layout_left:
    st.markdown("### 📂 File Upload")
    uploaded_file = st.file_uploader("Drop xlsx file here", type=["xlsx"], label_visibility="collapsed", key=f"xlsx_uploader_{st.session_state.uploader_id}")

valid_rows = pd.DataFrame()
total_records_count = 0
file_name_col, link_col = None, None

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        
        if len(df.columns) >= 2:
            file_name_col = df.columns[0]
            link_col = df.columns[1]
            
            valid_rows = df[df[link_col].notna() & df[link_col].astype(str).str.startswith('http')]
            total_records_count = len(valid_rows)
            
            if total_records_count > 0:
                with layout_left:
                    st.markdown("### 🔍 Excel Sheet Preview Data")
                    preview_df = df.copy()
                    preview_df.index = preview_df.index + 1
                    st.dataframe(preview_df.head(10), use_container_width=True, height=180)
                    
                with layout_right:
                    st.markdown("### 📊 File Uploaded")
                    st.markdown(f"""
                        <div class="metric-box">
                            <span style='font-size:13px; color:#6B7280; text-transform: uppercase; font-weight:bold;'>Folder Queue Load</span><br>
                            <span style='font-size:32px; font-weight:bold; color:#1E3A8A;'>{total_records_count}</span>
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ No valid HTTP links found in the second column.")
        else:
            st.error("❌ Column Validation Error: The sheet must contain at least 2 columns (Folder Name and Link).")
            
    except Exception as e:
        st.error(f"Failed parsing file: {e}")

with st.sidebar:
    st.write(" ")
    run_engine = st.button(
        "🚀 Start To Extract Images", 
        type="primary", 
        use_container_width=True,
        disabled=(uploaded_file is None or total_records_count == 0)
    )


# ====================================================
# WORKER FUNCTION: DOWNLOAD, EXTRACT, FLATTEN & CONVERT
# ====================================================
def process_rar_link(folder_name, rar_url, base_dir):
    folder_name = sanitize_filename(str(folder_name))
    
    # Auto-fix Bharath Cloud Links
    rar_url = str(rar_url).strip()
    if "/index.php/s/" in rar_url and not rar_url.endswith("/download"):
        rar_url = rar_url.rstrip("/") + "/download"
        
    target_dir = os.path.join(base_dir, folder_name)
    os.makedirs(target_dir, exist_ok=True)
    temp_rar_path = os.path.join(base_dir, f"temp_{folder_name}.rar")
    
    result = {"Folder Name": folder_name, "Status": "", "Extracted_Count": 0}
    success_flag = False
    
    try:
        response = requests.get(rar_url, headers=HEADERS, timeout=60, stream=True)
        content_type = response.headers.get('Content-Type', '')
        
        if 'text/html' in content_type:
            result["Status"] = "Failed: Link returned a webpage, not a file."
            return result, False
            
        if response.status_code == 200:
            with open(temp_rar_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    
            # 1. Extract RAR
            with rarfile.RarFile(temp_rar_path) as rf:
                rf.extractall(path=target_dir)
                
            # 2. Flatten subfolders
            for root, dirs, files in os.walk(target_dir):
                if root != target_dir: 
                    for file in files:
                        src_path = os.path.join(root, file)
                        dst_path = os.path.join(target_dir, file)
                        shutil.move(src_path, dst_path)
                        
            # 3. Clean up empty subfolders
            for root, dirs, files in os.walk(target_dir, topdown=False):
                if root != target_dir:
                    try: os.rmdir(root)
                    except OSError: pass
                        
            # 4. Convert .jfif to .jpg
            for filename in os.listdir(target_dir):
                if filename.lower().endswith(".jfif"):
                    old_path = os.path.join(target_dir, filename)
                    new_filename = filename[:-5] + ".jpg"
                    new_path = os.path.join(target_dir, new_filename)
                    os.rename(old_path, new_path)
                
            extracted_files = sum(len(files) for _, _, files in os.walk(target_dir))
            result["Extracted_Count"] = extracted_files
            result["Status"] = f"Success ({extracted_files} files extracted)"
            success_flag = True
            
        else:
            result["Status"] = f"Download Failed (HTTP {response.status_code})"
            
    except Exception as e:
        result["Status"] = f"Error: {str(e)}"
        
    finally:
        if os.path.exists(temp_rar_path):
            os.remove(temp_rar_path)
            
    return result, success_flag

# --- EXECUTION ENGINE ---
if run_engine and uploaded_file is not None:
    reset_metrics()
    
    working_dir = tempfile.mkdtemp(prefix="Manba_run_")
    output_package_dir = os.path.join(working_dir, "extracted_reports")
    os.makedirs(output_package_dir, exist_ok=True)
    
    st.write("---")
    st.markdown("### ⚙️ Images Extracting...>>> Logs")
        
    engine_progressbar = st.progress(0)
    percentage_text = st.empty()
    
    report_data = []
    completed_tasks = 0

    with st.status("Images extracting in pipelines....Please Wait....", expanded=True) as operation_context:
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_loan = {
                executor.submit(
                    process_rar_link, row[file_name_col], row[link_col], output_package_dir
                ): row[file_name_col] for _, row in valid_rows.iterrows()
            }
            
            for future in as_completed(future_to_loan):
                loan_id = future_to_loan[future]
                try:
                    res, is_success = future.result()
                    completed_tasks += 1
                    
                    st.session_state.total_extracted_items += res["Extracted_Count"]
                    if is_success and res["Extracted_Count"] > 0:
                        st.session_state.success_count += 1
                    else:
                        st.session_state.failure_count += 1
                        
                    report_data.append(res)
                    operation_context.write(f"✅ **Done ({completed_tasks}/{total_records_count}):** `{res['Folder Name']}` -> {res['Status']}")
                except Exception as e:
                    completed_tasks += 1
                    st.session_state.failure_count += 1
                    operation_context.write(f"🔴 **Error on {loan_id}:** {e}")
                    
                ratio = completed_tasks / total_records_count
                percentage_text.markdown(f"**📊 Images Extraction Progress: {int(ratio * 100)}%**")
                engine_progressbar.progress(ratio)
                gc.collect()

        operation_context.update(label="🚀 Images Extracted Successfully!", state="complete", expanded=False)

    if report_data:
        report_df = pd.DataFrame(report_data)
        st.session_state.final_df_summary = report_df
        
        report_csv_path = os.path.join(output_package_dir, "Operational_Reconciliation_Report.csv")
        report_df.to_csv(report_csv_path, index=False)
        
        master_delivery_zip = os.path.join(working_dir, "Manba_Master_Package.zip")
        with zipfile.ZipFile(master_delivery_zip, 'w', zipfile.ZIP_DEFLATED) as master_zip:
            for root, dirs, files in os.walk(output_package_dir):
                for file in files:
                    file_full_path = os.path.join(root, file)
                    relative_archive_path = os.path.relpath(file_full_path, output_package_dir)
                    master_zip.write(file_full_path, relative_archive_path)
                    
        st.session_state.zip_path = master_delivery_zip
        st.session_state.processing_complete = True
        st.balloons()

# --- RESULTS & DOWNLOAD RENDERING ---
if st.session_state.processing_complete and st.session_state.zip_path and os.path.exists(st.session_state.zip_path):
    st.write("---")
    st.markdown("### 🏁 Images Extracting Status In Details")
    
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #10B981;"><span class="card-title">Extracted Folders</span><div class="card-value" style="color:#10B981;">{st.session_state.success_count}</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #EF4444;"><span class="card-title">Failed Folders</span><div class="card-value" style="color:#EF4444;">{st.session_state.failure_count}</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="premium-card" style="border-bottom: 5px solid #F59E0B;"><span class="card-title">Total Extracted Images</span><div class="card-value" style="color:#F59E0B;">{st.session_state.total_extracted_items}</div></div>', unsafe_allow_html=True)
        
    tab_audit, tab_visuals = st.tabs(["📋 Detailed Extracting Report", "📊 Graphical Report"])
    
    with tab_audit:
        if st.session_state.final_df_summary is not None:
            st.markdown("##### 🔍 Real-Time Extracting Log")
            st.dataframe(st.session_state.final_df_summary, use_container_width=True, hide_index=True)
            
    with tab_visuals:
        st.markdown("##### 🚀 Images Extracting Ratio")
        health_chart_data = pd.DataFrame({"Status": ["Processed Successfully", "Failed"], "Values": [st.session_state.success_count, st.session_state.failure_count]})
        st.data_editor(
            health_chart_data,
            column_config={"Values": st.column_config.ProgressColumn("Ratio Scale", help="Clean output delivery distribution index", format="%d", min_value=0, max_value=total_records_count if total_records_count > 0 else 10)},
            hide_index=True, use_container_width=True, key="unified_visual_editor"
        )
        
    st.write(" ")
    with open(st.session_state.zip_path, "rb") as file_pointer:
        time_stamp = datetime.now(INDIAN_TZ).strftime('%d-%m-%Y_%H-%M-%S')
        custom_zip_filename = f"Manba_Export_{time_stamp}.zip"
            
        st.download_button(
            label="📥 DOWNLOAD IMAGES & REPORT",
            data=file_pointer,
            file_name=custom_zip_filename,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

st.markdown('<div class="footer-text">🛠️ This Images Tool Created by <b>Pawan Pandey</b></div>', unsafe_allow_html=True)