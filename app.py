import streamlit as st
import pandas as pd
import zipfile
import os
import warnings
import tempfile
import io

# Suppress annoying warnings
warnings.filterwarnings("ignore")

def read_mystery_file(filepath):
    """Tries every common 'disguised' Excel format until one works."""
    try: return pd.read_excel(filepath, engine='openpyxl')
    except Exception: pass
    
    try: return pd.read_excel(filepath, engine='xlrd')
    except Exception: pass
    
    try: return pd.read_csv(filepath, sep=None, engine='python', encoding='cp1252')
    except Exception: pass
    
    try: return pd.read_csv(filepath, sep='\t', encoding='utf-16')
    except Exception: pass

    try:
        dfs = pd.read_html(filepath)
        if dfs: return dfs[0]  
    except Exception: pass
    
    return None

@st.dialog("Data Warning") # Add the decorator
def show_duplicate_warning():
    st.write("⚠️ Duplicate, Null, or Conflicting entries were found in your submission.")
    st.write("Please review your data before proceeding.")
    
def apply_cdisc_mapping(df):
    """Applies the custom CDISC column mapping if the expected headers are present."""
    expected_header = ["Source Format Name", "Source Format Value", "CDISC Submission Value"]
    
    if all(col in df.columns for col in expected_header):
        output_df = pd.DataFrame(index=df.index)
        
        output_df["Source"] = "Custom"
        output_df["Codelist Name"] = df["Source Format Name"]
        output_df["Code"] = df["Source Format Value"]
        output_df["Decode/Label"] = df["CDISC Submission Value"]
        
        output_df["CT Name"] = ""
        output_df["CT Type"] = ""
        output_df["CT Code"] = ""
        output_df["CT Value"] = ""
        
        final_output_columns = [
            "Source", "Codelist Name", "Code", "Decode/Label", 
            "CT Name", "CT Type", "CT Code", "CT Value"
        ]
        return output_df[final_output_columns]
    
    return df

# --- UI CONFIGURATION ---
st.set_page_config(page_title="Codelist Combiner", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# --- ADVANCED UI HACKS ---
st.markdown("""
    <style>
    .stAppDeployButton {display:none;}
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Primary Action Buttons (Process Data) */
    div.stButton > button[kind="primary"] {
        background-color: #007BFF !important; 
        color: white !important;
        border: none !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #0056b3 !important; 
    }
    
    /* ✨ UPGRADED: RED DELETE/PURGE BUTTONS inside expanders */
    div[data-testid="stExpander"] button[kind="primary"] {
        background: linear-gradient(135deg, #FF4B4B 0%, #FF2E2E 100%) !important; 
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 10px rgba(255, 75, 75, 0.25) !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        white-space: nowrap !important; /* <--- ADD THIS LINE */
    }
    div[data-testid="stExpander"] button[kind="primary"]:hover {
        background: linear-gradient(135deg, #D32F2F 0%, #B71C1C 100%) !important;
        box-shadow: 0 6px 15px rgba(255, 75, 75, 0.4) !important;
        transform: translateY(-2px) !important;
    }

    /* ✨ UPGRADED: INSERT ROW BUTTON (General Secondary Buttons) */
    div.stButton > button[kind="secondary"] {
        background-color: #F8F9FA !important;
        color: #4F46E5 !important;
        border: 1px solid #4F46E5 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.02) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        background-color: #EEEDFF !important;
        box-shadow: 0 6px 12px rgba(79, 70, 229, 0.15) !important;
        transform: translateY(-2px) !important;
        border-style: solid !important;
    }

    /* Download Buttons */
    [data-testid="stDownloadButton"] button[kind="secondary"] {
        background-color: #107C41 !important; 
        color: white !important;
        border: none !important;
        border-style: none !important;
        border-radius: 6px !important;
        box-shadow: 0 4px 6px rgba(16, 124, 65, 0.25) !important;
    }
    [data-testid="stDownloadButton"] button[kind="secondary"]:hover {
        background-color: #0b5e31 !important;
        transform: translateY(-1px) !important;
    }

    [data-testid="stDownloadButton"] button[kind="primary"] {
        background-color: #4F46E5 !important; 
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        box-shadow: 0 4px 6px rgba(79, 70, 229, 0.25) !important; 
        transition: all 0.2s ease-in-out !important;
    }
    [data-testid="stDownloadButton"] button[kind="primary"]:hover {
        background-color: #4338CA !important;
        box-shadow: 0 6px 10px rgba(79, 70, 229, 0.4) !important;
        transform: translateY(-1px); 
    }

    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }

    [data-testid="stElementToolbar"] {
        opacity: 1 !important;
        visibility: visible !important;
        top: -35px !important;
    }
    
    
    /* HIDE THE '+' (ADD ROW) ICON IN THE DATA EDITOR TOOLBAR */
    [data-testid="stElementToolbar"] button[title="Add row"],
    [data-testid="stElementToolbar"] button[aria-label="Add row"],
    [data-testid="stElementToolbar"] button[title="Download as CSV"],
    [data-testid="stElementToolbar"] button[aria-label="Download as CSV"] {
        display: none !important;
    }

    [data-testid="stMetric"] {
        background-color: rgba(128, 128, 128, 0.05);
        padding: 15px;
        border-radius: 8px;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    </style>
""", unsafe_allow_html=True)


# --- SESSION STATE ---
if 'combined_df' not in st.session_state:
    st.session_state.combined_df = None
if 'file_count' not in st.session_state:
    st.session_state.file_count = 0
if 'failed_files' not in st.session_state:
    st.session_state.failed_files = []
if 'base_filename' not in st.session_state:
    st.session_state.base_filename = "Combined_Output"


# --- SIDEBAR ---
with st.sidebar:
    
    
    if st.session_state.combined_df is not None:
        st.warning("⚠️ **Note:** For CODELIST, download as EXCEL (.xlsx) for formatting purposes.")
        st.subheader("📥 Export Master File")
        
        clean_name = str(st.session_state.get('base_filename', 'Combined_Output')).strip()
        clean_name = os.path.splitext(clean_name)[0]  
        clean_name = "".join(c for c in clean_name if c.isalnum() or c in (' ', '_', '-')).rstrip()
        if clean_name == "":
            clean_name = "Combined_Output"
        
        # PREP EXCEL
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            st.session_state.combined_df.to_excel(writer, index=False)
        
        st.download_button(
            label="💾 Download Excel (.xlsx)", 
            data=output_buffer.getvalue(), 
            file_name=f"{clean_name}.xlsx", 
            type="secondary", 
            use_container_width=True
        )

        # PREP CSV
        csv_data = st.session_state.combined_df.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label="📄 Download CSV (.csv)", 
            data=csv_data, 
            file_name=f"{clean_name}.csv", 
            type="primary",
            use_container_width=True
        )
        
        st.divider()
        if st.button("🔄 Reset & Clear Data", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    else:
        
        st.title("📖 Quick Guide")
        st.divider()
        
        # Clean, modern step-by-step instructions
        st.markdown("""
        ### 
        
        📥 **1. Upload** Drop your `.zip` (UAT to PROD) or `.csv` (SAM to Platform).
        
        ⚙️ **2. Process** Click **Process Data** to extract, map, and stitch files.
        
        🧹 **3. Clean** Use the auto-saving data editor to fix or delete NULLs & Duplicates.
        
        💾 **4. Finalize** Review the master table and download in your preferred format.
        """)



# --- MAIN APP VIEW ---
st.title("Codelist Combiner")
st.divider()

if st.session_state.combined_df is None:
    st.subheader("Step 1: Upload Your Data")
    
    uploaded_file = st.file_uploader("Drop your ZIP archive or CSV file here", type=["zip", "csv"])

    if uploaded_file is not None:
        if st.button("🚀 Process Data", type="primary", use_container_width=True):
            
            st.session_state.base_filename = os.path.splitext(uploaded_file.name)[0]
            
            with tempfile.TemporaryDirectory() as temp_dir:
                with st.status("⚙️ Unpacking and processing...", expanded=True) as status:
                    
                    if uploaded_file.name.lower().endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(uploaded_file.read()), 'r') as zip_ref:
                            zip_ref.extractall(temp_dir)
                    else:
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    
                    all_files = []
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            if file.lower().endswith((".xlsx", ".xls", ".csv")) and not file.startswith("~$"):
                                all_files.append(os.path.join(root, file))
                    
                    if not all_files:
                        status.update(label="No valid files found to process.", state="error")
                        st.stop()

                    df_list = []
                    failed = []
                    prog = st.progress(0)
                    
                    for i, filename in enumerate(all_files):
                        df = read_mystery_file(filename)
                        if df is not None:
                            df_list.append(df)
                        else:
                            failed.append(os.path.basename(filename))
                        prog.progress((i + 1) / len(all_files))
                    
                    if df_list:
                        raw_combined = pd.concat(df_list, ignore_index=True)
                        mapped_combined = apply_cdisc_mapping(raw_combined)

                        st.session_state.combined_df = mapped_combined
                        st.session_state.file_count = len(df_list)
                        st.session_state.failed_files = failed
                        status.update(label="Stitching & Mapping Complete!", state="complete")
                        st.rerun() 
else:
    # --- SUCCESS STATE ---
    st.success("Files processed successfully! Manage your data below.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Files Combined", st.session_state.file_count)
    col2.metric("Total Rows", f"{len(st.session_state.combined_df):,}")
    col3.metric("Total Columns", len(st.session_state.combined_df.columns))
    
    if st.session_state.failed_files:
        with st.expander("⚠️ Failed to read some files"):
            for f in st.session_state.failed_files: st.markdown(f"- {f}")
            
    st.divider()

    # --- VALIDATION CHECKS (ACTIONABLE & AUTO-SAVING) ---
    df = st.session_state.combined_df
    check_cols = ["Source", "Codelist Name", "Code", "Decode/Label"]
    
    if all(col in df.columns for col in check_cols):
        # Calculate masks once
        null_mask = df[check_cols].isnull().any(axis=1)
        duplicate_mask = df.duplicated(subset=check_cols, keep=False)
        invalid_mask = df.groupby(['Codelist Name', 'Code'], dropna=False)['Decode/Label'].transform('nunique') > 1
        
        # Trigger the warning dialog if any issues are found
        if null_mask.any() or duplicate_mask.any() or invalid_mask.any():
            show_duplicate_warning()

        # 1. Actionable Null Check (Auto-Saves)
        if null_mask.any():
            with st.expander("🔍 Fix Missing Values (Auto-saves)", expanded=True):
                st.caption("Manual Adjustment: Edit values or use the trash icon on the left to delete rows. Saves automatically.")
                
                null_subset = df[null_mask]
                edited_nulls = st.data_editor(
                    null_subset, 
                    key="null_editor",
                    use_container_width=True, 
                    num_rows="dynamic",
                    hide_index=False
                )
                
                if not edited_nulls.equals(null_subset):
                    non_nulls = df[~null_mask]
                    st.session_state.combined_df = pd.concat([non_nulls, edited_nulls], ignore_index=True)
                    st.rerun()

        # 2. Actionable Duplicate Check (Auto-Saves)
        if duplicate_mask.any():
            with st.expander("🔍 Manage Duplicates (Auto-saves)", expanded=True):
                st.caption("Manual Adjustment: Edit values or use the trash icon on the left to delete rows. Saves automatically.")
                
                dupe_subset = df[duplicate_mask].sort_values(by=check_cols)
                edited_dupes = st.data_editor(
                    dupe_subset,
                    key="dupe_editor",
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=False
                )
                
                st.markdown("---")
                
                d_col1, d_col2 = st.columns([2, 3])
                if d_col1.button("🗑️ Remove Duplicates", type="primary", key="purge_dupes"):
                    st.session_state.combined_df = df.drop_duplicates(subset=check_cols, keep='first')
                    st.rerun()
                    
                if not edited_dupes.equals(dupe_subset):
                    non_dupes = df[~duplicate_mask]
                    st.session_state.combined_df = pd.concat([non_dupes, edited_dupes], ignore_index=True)
                    st.rerun()
                    
        # 3. Actionable Invalid Data Check (Conflicting Labels)
        if invalid_mask.any():
            with st.expander("⚠️ Manage Invalid Data (Conflicting Labels)", expanded=True):
                st.caption("These rows share the same **Codelist Name** and **Code**, but have conflicting **Decode/Labels**. Fix or delete rows below. Saves automatically.")
                
                invalid_subset = df[invalid_mask].sort_values(by=['Codelist Name', 'Code'])
                edited_invalid = st.data_editor(
                    invalid_subset,
                    key="invalid_editor",
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=False
                )
                
                if not edited_invalid.equals(invalid_subset):
                    valid_data = df[~invalid_mask]
                    st.session_state.combined_df = pd.concat([valid_data, edited_invalid], ignore_index=True)
                    st.rerun()

    st.subheader("Step 2: Final Review & Formatting")
    
    col_acc1, col_acc2 = st.columns([2, 3])
    with col_acc1:
        # Render the styled "Insert Row" button
        if st.button("➕ Insert Row", use_container_width=True):
            empty_row = pd.DataFrame([[None] * len(df.columns)], columns=df.columns)
            st.session_state.combined_df = pd.concat([empty_row, st.session_state.combined_df], ignore_index=True)
            st.rerun()

    st.session_state.combined_df = st.data_editor(
        st.session_state.combined_df, 
        use_container_width=True, 
        hide_index=False, 
        num_rows="dynamic",
        height=600
    )
#python3 -m streamlit run app.py