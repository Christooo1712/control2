import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import shutil
from datetime import datetime
import io

# Page configuration
st.set_page_config(
    page_title="IRCS2 Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<p class="main-header">📊 IRCS2 Report Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Internal Risk Control System - Data Validation & Reconciliation</p>', unsafe_allow_html=True)

# Sidebar
st.sidebar.image("https://via.placeholder.com/200x80/1f77b4/ffffff?text=IRCS2", use_container_width=True)
st.sidebar.markdown("## 📁 Upload Configuration")

# Upload Input Sheet
uploaded_input_sheet = st.sidebar.file_uploader(
    "Upload Input Sheet (Excel)",
    type=['xlsx'],
    help="Upload the Input Sheet_UAT.xlsx file that contains all paths and configurations"
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
### ℹ️ About
This tool automates the IRCS2 data validation and reconciliation process.

**What you need:**
- Input Sheet Excel file containing:
  - PATH INPUT sheet with all file paths
  - Code Library data
  - All required configurations

**Output:**
- Comprehensive Excel report with validation results
""")

# Main content
if uploaded_input_sheet is not None:
    try:
        # Save uploaded file to temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_input_sheet.getvalue())
            input_sheet_path = tmp_file.name
        
        # Import configuration module
        import IRCS2_input as input_config
        
        # Load configuration from Input Sheet
        with st.spinner("📖 Reading Input Sheet..."):
            path_map = input_config.load_config_from_excel(input_sheet_path)
        
        st.success("✅ Input Sheet loaded successfully!")
        
        # Display configuration in tabs
        tab1, tab2, tab3 = st.tabs(["📋 Configuration Summary", "📂 File Paths", "▶️ Run Process"])
        
        # Tab 1: Configuration Summary
        with tab1:
            st.markdown("### 📊 Reporting Configuration")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                reporting_month = path_map.get('Reporting Month')
                if reporting_month:
                    month_name = datetime(2025, int(reporting_month), 1).strftime('%B')
                    st.metric("Reporting Month", f"{month_name} ({reporting_month})")
                else:
                    st.metric("Reporting Month", "Not set")
            
            with col2:
                financial_year = path_map.get('Financial Year')
                if financial_year:
                    st.metric("Financial Year", int(financial_year))
                else:
                    st.metric("Financial Year", "Not set")
            
            with col3:
                output_filename = path_map.get('Output filename', 'IRCS2_Output')
                st.metric("Output Filename", f"{output_filename}.xlsx")
            
            st.markdown("---")
            st.markdown("### 📁 Input Files Configuration")
            
            # Create dataframe of all paths
            config_data = []
            for category, path in path_map.items():
                if category not in ['Reporting Month', 'Financial Year', 'Output filename']:
                    file_exists = "✅" if path and os.path.exists(path) else "❌"
                    config_data.append({
                        'Category': category,
                        'Path': path if path else 'Not configured',
                        'Exists': file_exists
                    })
            
            if config_data:
                config_df = pd.DataFrame(config_data)
                st.dataframe(config_df, use_container_width=True, hide_index=True)
                
                # Check if all files exist
                missing_files = [item['Category'] for item in config_data if item['Exists'] == '❌']
                if missing_files:
                    st.warning(f"⚠️ Warning: {len(missing_files)} file(s) not found on specified paths")
                    with st.expander("Show missing files"):
                        for file in missing_files:
                            st.write(f"- {file}")
                else:
                    st.success("✅ All configured files are accessible")
        
        # Tab 2: File Paths
        with tab2:
            st.markdown("### 📂 Complete Path Configuration")
            
            # Display all paths in an organized way
            path_df = pd.DataFrame(list(path_map.items()), columns=['Parameter', 'Value'])
            st.dataframe(path_df, use_container_width=True, hide_index=True)
            
            # Export configuration
            csv = path_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Configuration as CSV",
                data=csv,
                file_name="ircs2_configuration.csv",
                mime="text/csv",
            )
        
        # Tab 3: Run Process
        with tab3:
            st.markdown("### ▶️ Execute IRCS2 Process")
            
            st.info("""
            **Before running, please ensure:**
            1. All file paths in the Input Sheet are correct
            2. All required data files are accessible
            3. You have write permissions to the output directory
            """)
            
            # Check if critical files are configured
            critical_files = ['DV_AZTRAD', 'DV_AZUL', 'IT_AZTRAD', 'IT_AZUL', 'SUMMARY']
            missing_critical = [f for f in critical_files if not path_map.get(f)]
            
            if missing_critical:
                st.error(f"❌ Critical files not configured: {', '.join(missing_critical)}")
                st.stop()
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
                run_button = st.button(
                    "🚀 Run IRCS2 Process",
                    type="primary",
                    use_container_width=True
                )
            
            with col2:
                if st.button("🔄 Reload Config", use_container_width=True):
                    st.rerun()
            
            if run_button:
                # Progress tracking
                progress_bar = st.progress(0, text="Initializing...")
                status_text = st.empty()
                
                try:
                    # Step 1: Set paths
                    status_text.text("⚙️ Setting up configuration...")
                    progress_bar.progress(10, text="Setting up configuration...")
                    input_config.set_paths(path_map)
                    
                    # Step 2: Import modules
                    status_text.text("📦 Loading modules...")
                    progress_bar.progress(20, text="Loading modules...")
                    
                    # Clear any previously imported modules
                    modules_to_reload = ['Syntax.UL', 'Syntax.trad', 'Syntax.lookupvalue', 'IRCS2_program']
                    for module_name in modules_to_reload:
                        if module_name in sys.modules:
                            del sys.modules[module_name]
                    
                    # Step 3: Run processing
                    status_text.text("🔄 Processing data... This may take several minutes...")
                    progress_bar.progress(30, text="Processing data...")
                    
                    # Import and run the main program
                    import IRCS2_program
                    
                    progress_bar.progress(100, text="✅ Processing complete!")
                    status_text.text("")
                    
                    st.success("✅ IRCS2 processing completed successfully!")
                    st.balloons()
                    
                    # Step 4: Provide download
                    st.markdown("---")
                    st.markdown("### 📥 Download Results")
                    
                    output_path = input_config.xlsx_output
                    
                    if output_path and os.path.exists(output_path):
                        with open(output_path, 'rb') as f:
                            output_bytes = f.read()
                        
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.download_button(
                                label="📥 Download Excel Report",
                                data=output_bytes,
                                file_name=os.path.basename(output_path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        with col2:
                            file_size = len(output_bytes) / (1024 * 1024)  # MB
                            st.metric("File Size", f"{file_size:.2f} MB")
                        
                        st.info(f"📊 Report saved: {os.path.basename(output_path)}")
                    else:
                        st.warning("⚠️ Output file not found. Process may have failed.")
                
                except Exception as e:
                    progress_bar.empty()
                    status_text.empty()
                    
                    st.error("❌ An error occurred during processing")
                    
                    with st.expander("🔍 View Error Details", expanded=True):
                        st.code(str(e))
                        st.exception(e)
                    
                    st.markdown("### 💡 Troubleshooting Tips:")
                    st.markdown("""
                    1. Check if all file paths in the Input Sheet are correct
                    2. Ensure all data files are accessible
                    3. Verify file formats match expectations (CSV, Excel)
                    4. Check for data quality issues in source files
                    5. Review the error message above for specific issues
                    """)
        
        # Clean up temporary file
        if os.path.exists(input_sheet_path):
            os.unlink(input_sheet_path)
    
    except Exception as e:
        st.error(f"❌ Error loading Input Sheet: {str(e)}")
        with st.expander("🔍 View Error Details"):
            st.exception(e)

else:
    # Welcome screen when no file is uploaded
    st.info("👈 Please upload the Input Sheet Excel file from the sidebar to begin")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🚀 Quick Start Guide")
        st.markdown("""
        1. **Prepare Input Sheet**
           - Ensure PATH INPUT sheet contains all file paths
           - Verify Reporting Month and Financial Year
           - Check Output filename configuration
        
        2. **Upload File**
           - Click "Upload Input Sheet" in sidebar
           - Select your Input Sheet_UAT.xlsx file
        
        3. **Review Configuration**
           - Check all paths are correct
           - Verify all files are accessible
        
        4. **Run Process**
           - Go to "Run Process" tab
           - Click "Run IRCS2 Process"
           - Wait for completion
        
        5. **Download Report**
           - Download the generated Excel report
        """)
    
    with col2:
        st.markdown("### 📋 Input Sheet Requirements")
        st.markdown("""
        **Required Sheet: PATH INPUT**
        
        Must contain these categories:
        - Reporting Month
        - Financial Year
        - Output filename
        - DV_AZTRAD
        - DV_AZUL
        - IT_AZTRAD
        - IT_AZUL
        - SUMMARY
        - LGC_LGM_Campaign
        - BSI Attribusi
        - RESERVE_TRADCONV_RWNB_IFRS
        - RESERVE_TRADSHA_RWNB_IFRS
        - ACP Campaign
        - Code Library (optional, if using separate sheets)
        """)
        
        st.markdown("### 📊 Output Features")
        st.markdown("""
        - Summary Checking UL
        - Summary Checking TRAD
        - CONTROL_2_SUMMARY
        - Summary BSI
        - SUMMARY_CAMPAIGN
        - AZCP_CAMPAIGN
        - Automated formulas and formatting
        - Conditional formatting for errors
        """)

# Footer
st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("*IRCS2 Report Generator v1.0*")
with col2:
    st.markdown("*Internal Risk Control System*")
with col3:
    st.markdown(f"*Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}*")
