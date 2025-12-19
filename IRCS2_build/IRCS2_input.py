import pandas as pd
import os

# Global variables
CODE_LIBRARY_path = None
reporting_month = None
financial_year = None
DV_AZTRAD_path = None
DV_AZUL_path = None
IT_AZTRAD_path = None
IT_AZUL_path = None
SUMMARY_path = None
LGC_LGM_CAMPAIGN_path = None
BSI_ATTRIBUSI_path = None
TRADCONV_path = None
TRADSHA_path = None
acp_path = None
xlsx_output = None

def load_config_from_excel(uploaded_file):
    """Load configuration from uploaded Input Sheet Excel"""
    try:
        # Read PATH INPUT sheet
        input_df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name='PATH INPUT')
        
        # Create path mapping
        path_map = dict(zip(input_df['Category'], input_df['Path']))
        
        return path_map
    except Exception as e:
        raise Exception(f"Error reading Input Sheet: {str(e)}")

def set_paths(path_map):
    """Set global path variables from path_map dictionary"""
    global CODE_LIBRARY_path, reporting_month, financial_year
    global DV_AZTRAD_path, DV_AZUL_path, IT_AZTRAD_path, IT_AZUL_path
    global SUMMARY_path, LGC_LGM_CAMPAIGN_path, BSI_ATTRIBUSI_path
    global TRADCONV_path, TRADSHA_path, acp_path, xlsx_output
    
    # Set paths from mapping
    CODE_LIBRARY_path = path_map.get('Code Library')
    reporting_month = path_map.get('Reporting Month')
    financial_year = path_map.get('Financial Year')
    DV_AZTRAD_path = path_map.get('DV_AZTRAD')
    DV_AZUL_path = path_map.get('DV_AZUL')
    IT_AZTRAD_path = path_map.get('IT_AZTRAD')
    IT_AZUL_path = path_map.get('IT_AZUL')
    SUMMARY_path = path_map.get('SUMMARY')
    LGC_LGM_CAMPAIGN_path = path_map.get('LGC_LGM_Campaign')
    BSI_ATTRIBUSI_path = path_map.get('BSI Attribusi')
    TRADCONV_path = path_map.get('RESERVE_TRADCONV_RWNB_IFRS')
    TRADSHA_path = path_map.get('RESERVE_TRADSHA_RWNB_IFRS')
    acp_path = path_map.get('ACP Campaign')
    
    # Set output path
    xlsx_filename = path_map.get('Output filename')
    if DV_AZTRAD_path:
        output_dir = "\\".join([x for x in DV_AZTRAD_path.split('\\')]
                              [:len(DV_AZTRAD_path.split('\\')) - 1])
        xlsx_output = output_dir + "\\" + xlsx_filename + ".xlsx"
    else:
        xlsx_output = xlsx_filename + ".xlsx"
