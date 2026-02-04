# #!/usr/bin/env python3
# """
# FastAPI Application for Unified Digit TW (Two Wheeler) Processing
# ==================================================================
# This FastAPI app provides endpoints for processing TW patterns (1+5, COMP/SAOD, and TP)
# with automatic pattern detection.

# Endpoints:
#     POST /upload - Upload Excel file and process TW data
#     GET /health - Health check endpoint
#     GET / - API information

# Features:
# - Automatic worksheet detection and processing
# - Multiple TW pattern support (1+5, COMP/SAOD, TP)
# - File upload with validation
# - Comprehensive error handling
# - Returns processed data as JSON or downloadable Excel
# """

# from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Query
# from fastapi.responses import JSONResponse, FileResponse, StreamingResponse
# from fastapi.middleware.cors import CORSMiddleware
# from typing import List, Dict, Tuple, Optional
# import pandas as pd
# import io
# import os
# import re
# from datetime import datetime
# from pydantic import BaseModel
# import tempfile
# import traceback

# # ===============================================================================
# # FASTAPI APP SETUP
# # ===============================================================================

# app = FastAPI(
#     title="Digit TW Unified Processor API",
#     description="API for processing Two Wheeler insurance data with automatic pattern detection",
#     version="1.0.0"
# )

# # Add CORS middleware
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ===============================================================================
# # PYDANTIC MODELS
# # ===============================================================================

# class ProcessingResponse(BaseModel):
#     success: bool
#     message: str
#     pattern_detected: Optional[str] = None
#     total_records: int = 0
#     records: Optional[List[Dict]] = None
#     summary: Optional[Dict] = None
#     download_url: Optional[str] = None


# class HealthResponse(BaseModel):
#     status: str
#     version: str
#     timestamp: str


# # ===============================================================================
# # FORMULA DATA AND STATE MAPPING
# # ===============================================================================

# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
# ]

# STATE_MAPPING = {
#     "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS", "BIHAR": "BIHAR", "CHANDIGARH": "CHANDIGARH",
#     "DELHI": "DELHI", "GUJARAT": "GUJARAT", "GOA": "GOA", "HIMACHAL": "HIMACHAL PRADESH",
#     "J&K": "JAMMU AND KASHMIR", "JHARKHAND": "JHARKHAND", "MAHARASHTRA": "MAHARASHTRA",
#     "PUNJAB": "PUNJAB", "HARYANA": "HARYANA", "WEST BENGAL": "WEST BENGAL",
#     "UP": "UTTAR PRADESH", "UK": "UTTARAKHAND", "BR": "BIHAR", "DL": "DELHI", "GJ": "GUJARAT",
#     "MH": "MAHARASHTRA", "WB": "WEST BENGAL", "NCR": "DELHI NCR", "ROM": "REST OF MAHARASHTRA",
#     "RJ": "RAJASTHAN", "OR": "ODISHA", "AP": "ANDHRA PRADESH", "ASSAM": "ASSAM",
#     "KA": "KARNATAKA", "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA",
#     "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
#     "ROM1": "REST OF MAHARASHTRA", "ROM2": "REST OF MAHARASHTRA",
#     "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "GOOD RJ": "RAJASTHAN",
#     "NORTH BENGAL": "WEST BENGAL", "VIZAG": "ANDHRA PRADESH"
# }

# # ===============================================================================
# # CORE CALCULATION FUNCTIONS
# # ===============================================================================

# def safe_float(value) -> Optional[float]:
#     """Safely convert value to float, handling various edge cases."""
#     if pd.isna(value):
#         return None
#     val_str = str(value).strip().upper()
#     if val_str in ["D", "NA", "", "NAN", "NONE"]:
#         return None
#     try:
#         if isinstance(value, str):
#             value = value.strip().replace('%', '')
#         num = float(value)
#         if 0 < num < 1:
#             num = num * 100
#         return round(num, 2) if num > 0 else None
#     except:
#         return None


# def get_payin_category(payin: float) -> str:
#     """Categorize payin percentage into predefined ranges."""
#     if payin <= 20:
#         return "Payin Below 20%"
#     elif payin <= 30:
#         return "Payin 21% to 30%"
#     elif payin <= 50:
#         return "Payin 31% to 50%"
#     else:
#         return "Payin Above 50%"


# def extract_state(cluster_name: str) -> str:
#     """Extract state from cluster name using mapping."""
#     if pd.isna(cluster_name):
#         return "UNKNOWN"
#     cluster_upper = str(cluster_name).upper().strip()
    
#     for key, val in STATE_MAPPING.items():
#         if key in cluster_upper:
#             return val
    
#     parts = cluster_name.split('_')
#     if len(parts) > 0:
#         prefix = parts[0].upper()
#         if prefix in STATE_MAPPING:
#             return STATE_MAPPING[prefix]
    
#     return "REST OF INDIA"


# def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
#     """Get formula and calculate payout based on LOB, segment, policy type, and payin."""
#     segment_key = segment.upper()
    
#     if lob == "TW":
#         if policy_type == "TP":
#             segment_key = "TW TP"
#         elif segment == "1+5":
#             segment_key = "1+5"
#         else:
#             segment_key = "TW SAOD + COMP"
    
#     payin_category = get_payin_category(payin)
#     matching_rule = None
    
#     for rule in FORMULA_DATA:
#         if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
#             if rule["REMARKS"] == payin_category or rule["REMARKS"] == "NIL":
#                 matching_rule = rule
#                 break
    
#     if not matching_rule:
#         deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
#         return f"-{deduction}%", round(payin - deduction, 2)
    
#     formula = matching_rule["PO"]
    
#     if "% of Payin" in formula or "of Payin" in formula:
#         perc = float(formula.split("%")[0].strip().replace("Less ", ""))
#         if "Less" in formula:
#             return formula, round(payin - perc, 2)
#         else:
#             return formula, round(payin * perc / 100, 2)
#     elif formula.startswith("-") and "%" in formula:
#         ded = float(formula.replace("-", "").replace("%", ""))
#         return formula, round(payin - ded, 2)
#     else:
#         return "-2%", round(payin - 2, 2)


# def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
#     """Calculate payout with formula and rule explanation."""
#     if payin == 0:
#         return 0, "0% (No Payin)", "Payin is 0"
#     formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
#     return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# # ===============================================================================
# # PATTERN DETECTION
# # ===============================================================================

# class PatternTWDetector:
#     """Detects which TW pattern the sheet follows."""
    
#     @staticmethod
#     def detect_pattern(df: pd.DataFrame, sheet_name: str = "") -> str:
#         """Detect the pattern type based on sheet structure and name."""
#         sheet_upper = sheet_name.upper()
        
#         if "1+5" in sheet_upper or "1 5" in sheet_upper:
#             return '1plus5'
#         elif "SAOD" in sheet_upper or "COMP" in sheet_upper:
#             return 'comp_saod'
#         elif "TP" in sheet_upper or "SATP" in sheet_upper:
#             return 'tp'
        
#         df_str = df.head(20).to_string().upper()
        
#         if isinstance(df.columns, pd.Index):
#             columns = ' '.join([str(col).upper() for col in df.columns])
#         else:
#             first_rows = ' '.join([str(val).upper() for val in df.iloc[:5].values.flatten() if pd.notna(val)])
#             columns = first_rows
        
#         if "1+5" in columns or "1 5 CD2" in columns:
#             return '1plus5'
#         elif "YEAR CD2" in df_str or "1 YEAR" in df_str:
#             return 'comp_saod'
#         elif "SATP" in columns or ("TP" in columns and "CD2" in columns):
#             return 'tp'
        
#         num_cols = df.shape[1]
#         if num_cols <= 4:
#             return 'tp'
#         elif num_cols >= 6:
#             if "YEAR" in df_str:
#                 return 'comp_saod'
#             else:
#                 return '1plus5'
        
#         return 'tp'
    
#     @staticmethod
#     def detect_pattern_name(df: pd.DataFrame, sheet_name: str = "") -> str:
#         """Get a descriptive name for the detected pattern."""
#         pattern = PatternTWDetector.detect_pattern(df, sheet_name)
#         pattern_names = {
#             '1plus5': "TW 1+5 Pattern (90% of Payin)",
#             'comp_saod': "TW COMP/SAOD Pattern (Tiered Deductions)",
#             'tp': "TW TP Pattern (Third Party - Tiered Deductions)"
#         }
#         return pattern_names.get(pattern, "Unknown TW Pattern")

# # ===============================================================================
# # PATTERN PROCESSORS
# # ===============================================================================

# class OnePlusFiveProcessor:
#     """Process 1+5 pattern sheets for TW."""
    
#     @staticmethod
#     def process(df: pd.DataFrame, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None) -> List[Dict]:
#         """Process 1+5 pattern sheets."""
#         records = []
        
#         try:
#             # Find header row
#             header_row = None
#             for i in range(min(10, df.shape[0])):
#                 row_values = df.iloc[i].astype(str).str.upper()
#                 if "AGENCY" in row_values.values or "CLUSTER" in row_values.values or "MAKE" in row_values.values:
#                     header_row = i
#                     break
            
#             if header_row is None:
#                 header_row = 0
            
#             df.columns = df.iloc[header_row]
#             df = df.iloc[header_row + 1:].reset_index(drop=True)
#             df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
            
#             cluster_col = df.columns[0]
            
#             cd2_col = None
#             for i, col in enumerate(df.columns):
#                 col_upper = str(col).upper()
#                 if "CD2" in col_upper or ("1" in col_upper and "5" in col_upper):
#                     cd2_col = col
#                     break
            
#             if cd2_col is None:
#                 for idx in [4, 3, 5]:
#                     if idx < len(df.columns):
#                         cd2_col = df.columns[idx]
#                         break
            
#             segment_col = df.columns[1] if len(df.columns) > 1 else None
#             make_col = df.columns[1] if "MAKE" in str(df.columns[1]).upper() else None
            
#             lob_final = override_lob if override_enabled and override_lob else "TW"
#             segment_final = override_segment if override_enabled and override_segment else "1+5"
            
#             for idx, row in df.iterrows():
#                 cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
#                 if not cluster or cluster.upper() in ["", "TOTAL", "GRAND TOTAL", "AGENCY/PB CLUSTERS", "AGENCY", "CLUSTERS"]:
#                     continue
                
#                 payin = safe_float(row[cd2_col])
#                 if payin is None or payin <= 0:
#                     continue
                
#                 state = extract_state(cluster)
#                 payout = round(payin * 0.9, 2)
                
#                 original_segment = ""
#                 if segment_col:
#                     original_segment = str(row[segment_col]).strip() if pd.notna(row[segment_col]) else ""
                
#                 make = ""
#                 if make_col:
#                     make = str(row[make_col]).strip() if pd.notna(row[make_col]) else ""
                
#                 records.append({
#                     "State": state,
#                     "Location/Cluster": cluster,
#                     "Original Segment": original_segment,
#                     "Mapped Segment": segment_final,
#                     "LOB": lob_final,
#                     "Policy Type": "TP",
#                     "Status": "STP",
#                     "Payin (CD2)": f"{payin:.2f}%",
#                     "Payin Category": get_payin_category(payin),
#                     "Calculated Payout": f"{payout:.2f}%",
#                     "Formula Used": "90% of Payin",
#                     "Make": make,
#                     "Rule Explanation": "TW 1+5: Fixed 90% of Payin"
#                 })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in 1+5 processing: {e}")
#             traceback.print_exc()
#             return []


# class CompSaodTWProcessor:
#     """Process COMP/SAOD pattern sheets for TW."""
    
#     @staticmethod
#     def process(df: pd.DataFrame, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None) -> List[Dict]:
#         """Process COMP/SAOD pattern sheets with multiple year columns."""
#         records = []
        
#         try:
#             header_row = None
#             for i in range(min(10, df.shape[0])):
#                 row_str = " ".join(df.iloc[i].astype(str).str.upper())
#                 if ("CLUSTER" in row_str or "SEGMENT" in row_str) and "CD2" in row_str:
#                     header_row = i
#                     break
            
#             if header_row is None:
#                 header_row = 0
            
#             df.columns = df.iloc[header_row]
#             df = df.iloc[header_row + 1:].reset_index(drop=True)
#             df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
            
#             cluster_col = df.columns[0]
#             segment_col = df.columns[1] if len(df.columns) > 1 else None
            
#             cd2_columns = []
#             for col in df.columns:
#                 col_upper = str(col).upper()
#                 if "CD2" in col_upper:
#                     year_match = re.search(r'(\d+)\s*(?:YEAR|YR)', col_upper)
#                     if year_match:
#                         year_num = int(year_match.group(1))
#                         year_label = f"Year {year_num}"
#                     else:
#                         year_num = len(cd2_columns) + 1
#                         year_label = f"Year {year_num}"
                    
#                     cd2_columns.append((col, year_label, str(col)))
            
#             if not cd2_columns:
#                 return []
            
#             lob_final = override_lob if override_enabled and override_lob else "TW"
#             segment_final = override_segment if override_enabled and override_segment else "TW SAOD + COMP"
            
#             for idx, row in df.iterrows():
#                 cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
#                 if not cluster or cluster.upper() in ["CLUSTER", "TOTAL", "GRAND TOTAL", ""]:
#                     continue
                
#                 original_segment = ""
#                 if segment_col:
#                     original_segment = str(row[segment_col]).strip() if pd.notna(row[segment_col]) else ""
                
#                 state = extract_state(cluster)
                
#                 for cd2_col, year_label, orig_header in cd2_columns:
#                     payin = safe_float(row[cd2_col])
#                     if payin is None or payin == 0:
#                         continue
                    
#                     payout = round(payin * 0.9, 2)
                    
#                     records.append({
#                         "State": state,
#                         "Location/Cluster": cluster,
#                         "Original Segment": original_segment,
#                         "Mapped Segment": segment_final,
#                         "LOB": lob_final,
#                         "Policy Type": "SAOD",
#                         "Status": "STP",
#                         "Year": year_label,
#                         "Payin (CD2)": f"{payin:.2f}%",
#                         "Payin Category": get_payin_category(payin),
#                         "Calculated Payout": f"{payout:.2f}%",
#                         "Formula Used": "90% of Payin",
#                         "Rule Explanation": f"TW SAOD + COMP: {year_label}",
#                         "Remarks": orig_header
#                     })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in COMP/SAOD processing: {e}")
#             traceback.print_exc()
#             return []


# class TPProcessor:
#     """Process TP (Third Party) pattern sheets for TW."""
    
#     @staticmethod
#     def process(df: pd.DataFrame, sheet_name: str,
#                 override_enabled: bool = False,
#                 override_lob: str = None,
#                 override_segment: str = None,
#                 override_policy_type: str = None) -> List[Dict]:
#         """Process TP pattern sheets."""
#         records = []
        
#         try:
#             start_row = 0
#             for idx, row in df.iterrows():
#                 cell_value = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
#                 if any(keyword in cell_value for keyword in ["AGENCY", "CLUSTER", "NCR", "LOCATION"]):
#                     if cell_value not in ["AGENCY/PB CLUSTERS", "AGENCY", "CLUSTERS"]:
#                         start_row = idx
#                         break
#                     else:
#                         start_row = idx + 1
#                         break
            
#             cluster_col = 0
#             segment_col = 1 if df.shape[1] > 1 else None
#             cd2_col = 2 if df.shape[1] > 2 else df.shape[1] - 1
            
#             lob_final = override_lob if override_enabled and override_lob else "TW"
#             segment_final = override_segment if override_enabled and override_segment else "TW TP"
#             policy_final = override_policy_type if override_policy_type else "TP"
            
#             for idx in range(start_row, len(df)):
#                 row = df.iloc[idx]
                
#                 cluster = str(row.iloc[cluster_col]).strip() if pd.notna(row.iloc[cluster_col]) else ""
#                 if not cluster or cluster.upper() in ["TOTAL", "GRAND TOTAL", "AGENCY", "CLUSTERS"]:
#                     continue
                
#                 payin = safe_float(row.iloc[cd2_col])
#                 if payin is None:
#                     continue
                
#                 state = extract_state(cluster)
                
#                 original_segment = ""
#                 if segment_col is not None and segment_col < len(row):
#                     original_segment = str(row.iloc[segment_col]).strip() if pd.notna(row.iloc[segment_col]) else ""
                
#                 payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                
#                 records.append({
#                     "State": state,
#                     "Location/Cluster": cluster,
#                     "Original Segment": original_segment,
#                     "Mapped Segment": segment_final,
#                     "LOB": lob_final,
#                     "Policy Type": policy_final,
#                     "Status": "STP",
#                     "Payin (CD2)": f"{payin:.2f}%",
#                     "Payin Category": get_payin_category(payin),
#                     "Calculated Payout": f"{payout:.2f}%",
#                     "Formula Used": formula,
#                     "Rule Explanation": rule_exp
#                 })
            
#             return records
            
#         except Exception as e:
#             print(f"Error in TP processing: {e}")
#             traceback.print_exc()
#             return []

# # ===============================================================================
# # PATTERN DISPATCHER
# # ===============================================================================

# class PatternTWDispatcher:
#     """Main dispatcher that routes to appropriate TW pattern processor."""
    
#     PATTERN_PROCESSORS = {
#         '1plus5': OnePlusFiveProcessor,
#         'comp_saod': CompSaodTWProcessor,
#         'tp': TPProcessor
#     }
    
#     @staticmethod
#     def process_sheet(df: pd.DataFrame, sheet_name: str,
#                      override_enabled: bool = False,
#                      override_lob: str = None,
#                      override_segment: str = None,
#                      override_policy_type: str = None) -> List[Dict]:
#         """Main entry point for processing any TW sheet."""
#         pattern = PatternTWDetector.detect_pattern(df, sheet_name)
#         processor_class = PatternTWDispatcher.PATTERN_PROCESSORS.get(pattern, TPProcessor)
        
#         if pattern == 'tp':
#             records = processor_class.process(
#                 df, sheet_name, override_enabled,
#                 override_lob, override_segment, override_policy_type
#             )
#         else:
#             records = processor_class.process(
#                 df, sheet_name, override_enabled,
#                 override_lob, override_segment
#             )
        
#         return records

# # ===============================================================================
# # UTILITY FUNCTIONS
# # ===============================================================================

# def generate_summary(records: List[Dict]) -> Dict:
#     """Generate summary statistics for processed records."""
#     if not records:
#         return {}
    
#     summary = {
#         "total_records": len(records),
#         "states": {},
#         "policy_types": {},
#         "patterns": {},
#         "avg_payin": 0,
#         "avg_payout": 0
#     }
    
#     payins = []
#     payouts = []
    
#     for record in records:
#         # Count by state
#         state = record.get("State", "Unknown")
#         summary["states"][state] = summary["states"].get(state, 0) + 1
        
#         # Count by policy type
#         policy = record.get("Policy Type", "Unknown")
#         summary["policy_types"][policy] = summary["policy_types"].get(policy, 0) + 1
        
#         # Count by pattern
#         segment = record.get("Mapped Segment", "Unknown")
#         summary["patterns"][segment] = summary["patterns"].get(segment, 0) + 1
        
#         # Calculate averages
#         try:
#             payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
#             payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
#             if payin > 0:
#                 payins.append(payin)
#                 payouts.append(payout)
#         except:
#             pass
    
#     if payins:
#         summary["avg_payin"] = round(sum(payins) / len(payins), 2)
#         summary["avg_payout"] = round(sum(payouts) / len(payouts), 2)
    
#     return summary


# def get_sheet_names_from_file(file_path: str) -> List[str]:
#     """Get all sheet names from an Excel file."""
#     try:
#         xls = pd.ExcelFile(file_path)
#         return xls.sheet_names
#     except Exception as e:
#         raise HTTPException(status_code=400, detail=f"Error reading Excel file: {str(e)}")


# def find_best_worksheet(sheets: List[str]) -> Optional[str]:
#     """
#     Automatically find the best worksheet to process.
#     Priority: TW sheets, then first non-empty sheet.
#     """
#     if not sheets:
#         return None
    
#     # Priority keywords for TW sheets
#     tw_keywords = ["TW", "1+5", "TP", "SAOD", "COMP", "TWO WHEELER", "2W"]
    
#     # First pass: look for TW-related sheets
#     for sheet in sheets:
#         sheet_upper = sheet.upper()
#         for keyword in tw_keywords:
#             if keyword in sheet_upper:
#                 return sheet
    
#     # Second pass: avoid summary/total sheets
#     avoid_keywords = ["SUMMARY", "TOTAL", "INDEX", "MASTER", "REFERENCE"]
#     for sheet in sheets:
#         sheet_upper = sheet.upper()
#         if not any(avoid in sheet_upper for avoid in avoid_keywords):
#             return sheet
    
#     # Fallback: return first sheet
#     return sheets[0]


# # ===============================================================================
# # FASTAPI ENDPOINTS
# # ===============================================================================

# @app.get("/", response_model=dict)
# async def root():
#     """Root endpoint with API information."""
#     return {
#         "app": "Digit TW Unified Processor API",
#         "version": "1.0.0",
#         "description": "API for processing Two Wheeler insurance data with automatic pattern detection",
#         "endpoints": {
#             "POST /upload": "Upload and process Excel file",
#             "GET /health": "Health check",
#             "GET /": "This information"
#         },
#         "supported_patterns": [
#             "TW 1+5 Pattern (90% of Payin)",
#             "TW COMP/SAOD Pattern (Tiered Deductions)",
#             "TW TP Pattern (Third Party - Tiered Deductions)"
#         ]
#     }


# @app.get("/health", response_model=HealthResponse)
# async def health_check():
#     """Health check endpoint."""
#     return HealthResponse(
#         status="healthy",
#         version="1.0.0",
#         timestamp=datetime.now().isoformat()
#     )


# @app.post("/upload")
# async def upload_file(
#     file: UploadFile = File(..., description="Excel file to process"),
#     sheet_name: Optional[str] = Form(None, description="Specific sheet name to process (optional)"),
#     return_excel: bool = Form(False, description="Return downloadable Excel file"),
#     override_lob: Optional[str] = Form(None, description="Override LOB (default: TW)"),
#     override_segment: Optional[str] = Form(None, description="Override segment"),
#     override_policy_type: Optional[str] = Form(None, description="Override policy type")
# ):
#     """
#     Upload and process Excel file with automatic worksheet detection.
    
#     Parameters:
#     - file: Excel file (.xlsx, .xls)
#     - sheet_name: Optional specific sheet name. If not provided, auto-detects best sheet
#     - return_excel: If True, returns downloadable Excel file instead of JSON
#     - override_lob: Optional LOB override (default: TW)
#     - override_segment: Optional segment override
#     - override_policy_type: Optional policy type override
    
#     Returns:
#     - JSON with processed records and summary, or downloadable Excel file
#     """
    
#     # Validate file type
#     if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
#         raise HTTPException(
#             status_code=400,
#             detail="Invalid file type. Please upload an Excel file (.xlsx, .xls, .xlsm)"
#         )
    
#     # Create temporary file
#     temp_file = None
#     try:
#         # Save uploaded file to temporary location
#         with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
#             content = await file.read()
#             temp_file.write(content)
#             temp_file_path = temp_file.name
        
#         # Get all sheet names
#         sheets = get_sheet_names_from_file(temp_file_path)
        
#         if not sheets:
#             raise HTTPException(status_code=400, detail="No worksheets found in Excel file")
        
#         # Determine which sheet to process
#         if sheet_name:
#             # User specified a sheet
#             if sheet_name not in sheets:
#                 raise HTTPException(
#                     status_code=400,
#                     detail=f"Sheet '{sheet_name}' not found. Available sheets: {', '.join(sheets)}"
#                 )
#             selected_sheet = sheet_name
#         else:
#             # Auto-detect best sheet
#             selected_sheet = find_best_worksheet(sheets)
#             if not selected_sheet:
#                 raise HTTPException(status_code=400, detail="Could not determine worksheet to process")
        
#         print(f"Processing sheet: {selected_sheet}")
        
#         # Load the selected sheet
#         df = pd.read_excel(temp_file_path, sheet_name=selected_sheet, header=None)
        
#         # Check if override is enabled
#         override_enabled = bool(override_lob or override_segment or override_policy_type)
        
#         # Process the sheet
#         records = PatternTWDispatcher.process_sheet(
#             df, selected_sheet, override_enabled,
#             override_lob, override_segment, override_policy_type
#         )
        
#         if not records:
#             raise HTTPException(
#                 status_code=422,
#                 detail="No records extracted from the file. Please check the file structure."
#             )
        
#         # Detect pattern for response
#         pattern_name = PatternTWDetector.detect_pattern_name(df, selected_sheet)
        
#         # Generate summary
#         summary = generate_summary(records)
        
#         # Return Excel file if requested
#         if return_excel:
#             output_df = pd.DataFrame(records)
            
#             # Create Excel file in memory
#             output = io.BytesIO()
#             with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                 output_df.to_excel(writer, index=False, sheet_name='Processed')
#             output.seek(0)
            
#             # Generate filename
#             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#             filename = f"TW_Processed_{selected_sheet.replace(' ', '_')}_{timestamp}.xlsx"
            
#             return StreamingResponse(
#                 output,
#                 media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 headers={"Content-Disposition": f"attachment; filename={filename}"}
#             )
        
#         # Return JSON response
#         return JSONResponse(content={
#             "success": True,
#             "message": f"Successfully processed {len(records)} records from sheet '{selected_sheet}'",
#             "pattern_detected": pattern_name,
#             "sheet_processed": selected_sheet,
#             "available_sheets": sheets,
#             "total_records": len(records),
#             "summary": summary,
#             "records": records[:100] if len(records) > 100 else records,  # Limit to 100 for response size
#             "note": "Only first 100 records shown in response. Use return_excel=true to get all records."
#         })
    
#     except HTTPException:
#         raise
#     except Exception as e:
#         print(f"Error processing file: {e}")
#         traceback.print_exc()
#         raise HTTPException(
#             status_code=500,
#             detail=f"Error processing file: {str(e)}"
#         )
#     finally:
#         # Clean up temporary file
#         if temp_file and os.path.exists(temp_file_path):
#             try:
#                 os.unlink(temp_file_path)
#             except:
#                 pass


# @app.get("/sheets")
# async def get_sheets(file: UploadFile = File(..., description="Excel file")):
#     """
#     Get list of all sheets in an Excel file without processing.
#     Useful for determining which sheet to process.
#     """
#     if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
#         raise HTTPException(
#             status_code=400,
#             detail="Invalid file type. Please upload an Excel file (.xlsx, .xls, .xlsm)"
#         )
    
#     temp_file = None
#     try:
#         with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
#             content = await file.read()
#             temp_file.write(content)
#             temp_file_path = temp_file.name
        
#         sheets = get_sheet_names_from_file(temp_file_path)
#         recommended_sheet = find_best_worksheet(sheets)
        
#         return {
#             "success": True,
#             "total_sheets": len(sheets),
#             "sheets": sheets,
#             "recommended_sheet": recommended_sheet,
#             "message": f"Found {len(sheets)} sheet(s) in the file"
#         }
    
#     except Exception as e:
#         raise HTTPException(
#             status_code=500,
#             detail=f"Error reading sheets: {str(e)}"
#         )
#     finally:
#         if temp_file and os.path.exists(temp_file_path):
#             try:
#                 os.unlink(temp_file_path)
#             except:
#                 pass


# # ===============================================================================
# # RUN APPLICATION
# # ===============================================================================

# if __name__ == "__main__":
#     import uvicorn
#     uvicorn.run(app, host="0.0.0.0", port=8000)

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import io
import os
import re
from typing import List, Dict, Tuple, Optional
from datetime import datetime
import traceback
import tempfile

app = FastAPI(title="DIGIT TW Processor API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===============================================================================
# FORMULA DATA AND STATE MAPPING
# ===============================================================================

FORMULA_DATA = [
    {"LOB": "TW", "SEGMENT": "1+5", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
]

STATE_MAPPING = {
    "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS", "BIHAR": "BIHAR", "CHANDIGARH": "CHANDIGARH",
    "DELHI": "DELHI", "GUJARAT": "GUJARAT", "GOA": "GOA", "HIMACHAL": "HIMACHAL PRADESH",
    "J&K": "JAMMU AND KASHMIR", "JHARKHAND": "JHARKHAND", "MAHARASHTRA": "MAHARASHTRA",
    "PUNJAB": "PUNJAB", "HARYANA": "HARYANA", "WEST BENGAL": "WEST BENGAL",
    "UP": "UTTAR PRADESH", "UK": "UTTARAKHAND", "BR": "BIHAR", "DL": "DELHI", "GJ": "GUJARAT",
    "MH": "MAHARASHTRA", "WB": "WEST BENGAL", "NCR": "DELHI NCR", "ROM": "REST OF MAHARASHTRA",
    "RJ": "RAJASTHAN", "OR": "ODISHA", "AP": "ANDHRA PRADESH", "ASSAM": "ASSAM",
    "KA": "KARNATAKA", "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA",
    "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
    "ROM1": "REST OF MAHARASHTRA", "ROM2": "REST OF MAHARASHTRA",
    "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "GOOD RJ": "RAJASTHAN",
    "NORTH BENGAL": "WEST BENGAL", "VIZAG": "ANDHRA PRADESH"
}

# Store uploaded files temporarily
uploaded_files = {}

# ===============================================================================
# CORE CALCULATION FUNCTIONS
# ===============================================================================

def safe_float(value) -> Optional[float]:
    """Safely convert value to float, handling various edge cases."""
    if pd.isna(value):
        return None
    val_str = str(value).strip().upper()
    if val_str in ["D", "NA", "", "NAN", "NONE"]:
        return None
    try:
        if isinstance(value, str):
            value = value.strip().replace('%', '')
        num = float(value)
        if 0 < num < 1:
            num = num * 100
        return round(num, 2) if num > 0 else None
    except:
        return None


def get_payin_category(payin: float) -> str:
    """Categorize payin percentage into predefined ranges."""
    if payin <= 20:
        return "Payin Below 20%"
    elif payin <= 30:
        return "Payin 21% to 30%"
    elif payin <= 50:
        return "Payin 31% to 50%"
    else:
        return "Payin Above 50%"


def extract_state(cluster_name: str) -> str:
    """Extract state from cluster name using mapping."""
    if pd.isna(cluster_name):
        return "UNKNOWN"
    cluster_upper = str(cluster_name).upper().strip()
    
    for key, val in STATE_MAPPING.items():
        if key in cluster_upper:
            return val
    
    parts = cluster_name.split('_')
    if len(parts) > 0:
        prefix = parts[0].upper()
        if prefix in STATE_MAPPING:
            return STATE_MAPPING[prefix]
    
    return "REST OF INDIA"


def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[str, float]:
    """Get formula and calculate payout based on LOB, segment, policy type, and payin."""
    segment_key = segment.upper()
    
    if lob == "TW":
        if policy_type == "TP":
            segment_key = "TW TP"
        elif segment == "1+5":
            segment_key = "1+5"
        else:
            segment_key = "TW SAOD + COMP"
    
    payin_category = get_payin_category(payin)
    matching_rule = None
    
    for rule in FORMULA_DATA:
        if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
            if rule["REMARKS"] == payin_category or rule["REMARKS"] == "NIL":
                matching_rule = rule
                break
    
    if not matching_rule:
        deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
        return f"-{deduction}%", round(payin - deduction, 2)
    
    formula = matching_rule["PO"]
    
    if "% of Payin" in formula or "of Payin" in formula:
        perc = float(formula.split("%")[0].strip().replace("Less ", ""))
        if "Less" in formula:
            return formula, round(payin - perc, 2)
        else:
            return formula, round(payin * perc / 100, 2)
    elif formula.startswith("-") and "%" in formula:
        ded = float(formula.replace("-", "").replace("%", ""))
        return formula, round(payin - ded, 2)
    else:
        return "-2%", round(payin - 2, 2)


def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float) -> Tuple[float, str, str]:
    """Calculate payout with formula and rule explanation."""
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
    return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

# ===============================================================================
# PATTERN DETECTION
# ===============================================================================

class PatternTWDetector:
    """Detects which TW pattern the sheet follows."""
    
    @staticmethod
    def detect_pattern(df: pd.DataFrame, sheet_name: str = "") -> str:
        """Detect the pattern type based on sheet structure and name."""
        sheet_upper = sheet_name.upper()
        
        if "1+5" in sheet_upper or "1 5" in sheet_upper:
            return '1plus5'
        elif "SAOD" in sheet_upper or "COMP" in sheet_upper:
            return 'comp_saod'
        elif "TP" in sheet_upper or "SATP" in sheet_upper:
            return 'tp'
        
        df_str = df.head(20).to_string().upper()
        
        if isinstance(df.columns, pd.Index):
            columns = ' '.join([str(col).upper() for col in df.columns])
        else:
            first_rows = ' '.join([str(val).upper() for val in df.iloc[:5].values.flatten() if pd.notna(val)])
            columns = first_rows
        
        if "1+5" in columns or "1 5 CD2" in columns:
            return '1plus5'
        elif "YEAR CD2" in df_str or "1 YEAR" in df_str:
            return 'comp_saod'
        elif "SATP" in columns or ("TP" in columns and "CD2" in columns):
            return 'tp'
        
        num_cols = df.shape[1]
        if num_cols <= 4:
            return 'tp'
        elif num_cols >= 6:
            if "YEAR" in df_str:
                return 'comp_saod'
            else:
                return '1plus5'
        
        return 'tp'
    
    @staticmethod
    def detect_pattern_name(df: pd.DataFrame, sheet_name: str = "") -> str:
        """Get a descriptive name for the detected pattern."""
        pattern = PatternTWDetector.detect_pattern(df, sheet_name)
        pattern_names = {
            '1plus5': "TW 1+5 Pattern (90% of Payin)",
            'comp_saod': "TW COMP/SAOD Pattern (Tiered Deductions)",
            'tp': "TW TP Pattern (Third Party - Tiered Deductions)"
        }
        return pattern_names.get(pattern, "Unknown TW Pattern")

# ===============================================================================
# PATTERN PROCESSORS
# ===============================================================================

class OnePlusFiveProcessor:
    """Process 1+5 pattern sheets for TW."""
    
    @staticmethod
    def process(df: pd.DataFrame, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None) -> List[Dict]:
        """Process 1+5 pattern sheets."""
        records = []
        
        try:
            header_row = None
            for i in range(min(10, df.shape[0])):
                row_values = df.iloc[i].astype(str).str.upper()
                if "AGENCY" in row_values.values or "CLUSTER" in row_values.values or "MAKE" in row_values.values:
                    header_row = i
                    break
            
            if header_row is None:
                header_row = 0
            
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
            
            cluster_col = df.columns[0]
            
            cd2_col = None
            for i, col in enumerate(df.columns):
                col_upper = str(col).upper()
                if "CD2" in col_upper or ("1" in col_upper and "5" in col_upper):
                    cd2_col = col
                    break
            
            if cd2_col is None:
                for idx in [4, 3, 5]:
                    if idx < len(df.columns):
                        cd2_col = df.columns[idx]
                        break
            
            segment_col = df.columns[1] if len(df.columns) > 1 else None
            make_col = df.columns[1] if "MAKE" in str(df.columns[1]).upper() else None
            
            lob_final = override_lob if override_enabled and override_lob else "TW"
            segment_final = override_segment if override_enabled and override_segment else "1+5"
            
            for idx, row in df.iterrows():
                cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
                if not cluster or cluster.upper() in ["", "TOTAL", "GRAND TOTAL", "AGENCY/PB CLUSTERS", "AGENCY", "CLUSTERS"]:
                    continue
                
                payin = safe_float(row[cd2_col])
                if payin is None or payin <= 0:
                    continue
                
                state = extract_state(cluster)
                payout = round(payin * 0.9, 2)
                
                original_segment = ""
                if segment_col:
                    original_segment = str(row[segment_col]).strip() if pd.notna(row[segment_col]) else ""
                
                make = ""
                if make_col:
                    make = str(row[make_col]).strip() if pd.notna(row[make_col]) else ""
                
                records.append({
                    "State": state,
                    "Location/Cluster": cluster,
                    "Original Segment": original_segment,
                    "Mapped Segment": segment_final,
                    "LOB": lob_final,
                    "Policy Type": "TP",
                    "Status": "STP",
                    "Payin (CD2)": f"{payin:.2f}%",
                    "Payin Category": get_payin_category(payin),
                    "Calculated Payout": f"{payout:.2f}%",
                    "Formula Used": "90% of Payin",
                    "Make": make,
                    "Rule Explanation": "TW 1+5: Fixed 90% of Payin"
                })
            
            return records
            
        except Exception as e:
            print(f"Error in 1+5 processing: {e}")
            traceback.print_exc()
            return []


class CompSaodTWProcessor:
    """Process COMP/SAOD pattern sheets for TW."""
    
    @staticmethod
    def process(df: pd.DataFrame, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None) -> List[Dict]:
        """Process COMP/SAOD pattern sheets with multiple year columns."""
        records = []
        
        try:
            header_row = None
            for i in range(min(10, df.shape[0])):
                row_str = " ".join(df.iloc[i].astype(str).str.upper())
                if ("CLUSTER" in row_str or "SEGMENT" in row_str) and "CD2" in row_str:
                    header_row = i
                    break
            
            if header_row is None:
                header_row = 0
            
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
            
            cluster_col = df.columns[0]
            segment_col = df.columns[1] if len(df.columns) > 1 else None
            
            cd2_columns = []
            for col in df.columns:
                col_upper = str(col).upper()
                if "CD2" in col_upper:
                    year_match = re.search(r'(\d+)\s*(?:YEAR|YR)', col_upper)
                    if year_match:
                        year_num = int(year_match.group(1))
                        year_label = f"Year {year_num}"
                    else:
                        year_num = len(cd2_columns) + 1
                        year_label = f"Year {year_num}"
                    
                    cd2_columns.append((col, year_label, str(col)))
            
            if not cd2_columns:
                return []
            
            lob_final = override_lob if override_enabled and override_lob else "TW"
            segment_final = override_segment if override_enabled and override_segment else "TW SAOD + COMP"
            
            for idx, row in df.iterrows():
                cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
                if not cluster or cluster.upper() in ["CLUSTER", "TOTAL", "GRAND TOTAL", ""]:
                    continue
                
                original_segment = ""
                if segment_col:
                    original_segment = str(row[segment_col]).strip() if pd.notna(row[segment_col]) else ""
                
                state = extract_state(cluster)
                
                for cd2_col, year_label, orig_header in cd2_columns:
                    payin = safe_float(row[cd2_col])
                    if payin is None or payin == 0:
                        continue
                    
                    payout = round(payin * 0.9, 2)
                    
                    records.append({
                        "State": state,
                        "Location/Cluster": cluster,
                        "Original Segment": original_segment,
                        "Mapped Segment": segment_final,
                        "LOB": lob_final,
                        "Policy Type": "SAOD",
                        "Status": "STP",
                        "Year": year_label,
                        "Payin (CD2)": f"{payin:.2f}%",
                        "Payin Category": get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%",
                        "Formula Used": "90% of Payin",
                        "Rule Explanation": f"TW SAOD + COMP: {year_label}",
                        "Remarks": orig_header
                    })
            
            return records
            
        except Exception as e:
            print(f"Error in COMP/SAOD processing: {e}")
            traceback.print_exc()
            return []


class TPProcessor:
    """Process TP (Third Party) pattern sheets for TW."""
    
    @staticmethod
    def process(df: pd.DataFrame, sheet_name: str,
                override_enabled: bool = False,
                override_lob: str = None,
                override_segment: str = None,
                override_policy_type: str = None) -> List[Dict]:
        """Process TP pattern sheets."""
        records = []
        
        try:
            start_row = 0
            for idx, row in df.iterrows():
                cell_value = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
                if any(keyword in cell_value for keyword in ["AGENCY", "CLUSTER", "NCR", "LOCATION"]):
                    if cell_value not in ["AGENCY/PB CLUSTERS", "AGENCY", "CLUSTERS"]:
                        start_row = idx
                        break
                    else:
                        start_row = idx + 1
                        break
            
            cluster_col = 0
            segment_col = 1 if df.shape[1] > 1 else None
            cd2_col = 2 if df.shape[1] > 2 else df.shape[1] - 1
            
            lob_final = override_lob if override_enabled and override_lob else "TW"
            segment_final = override_segment if override_enabled and override_segment else "TW TP"
            policy_final = override_policy_type if override_policy_type else "TP"
            
            for idx in range(start_row, len(df)):
                row = df.iloc[idx]
                
                cluster = str(row.iloc[cluster_col]).strip() if pd.notna(row.iloc[cluster_col]) else ""
                if not cluster or cluster.upper() in ["TOTAL", "GRAND TOTAL", "AGENCY", "CLUSTERS"]:
                    continue
                
                payin = safe_float(row.iloc[cd2_col])
                if payin is None:
                    continue
                
                state = extract_state(cluster)
                
                original_segment = ""
                if segment_col is not None and segment_col < len(row):
                    original_segment = str(row.iloc[segment_col]).strip() if pd.notna(row.iloc[segment_col]) else ""
                
                payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_final, payin)
                
                records.append({
                    "State": state,
                    "Location/Cluster": cluster,
                    "Original Segment": original_segment,
                    "Mapped Segment": segment_final,
                    "LOB": lob_final,
                    "Policy Type": policy_final,
                    "Status": "STP",
                    "Payin (CD2)": f"{payin:.2f}%",
                    "Payin Category": get_payin_category(payin),
                    "Calculated Payout": f"{payout:.2f}%",
                    "Formula Used": formula,
                    "Rule Explanation": rule_exp
                })
            
            return records
            
        except Exception as e:
            print(f"Error in TP processing: {e}")
            traceback.print_exc()
            return []

# ===============================================================================
# PATTERN DISPATCHER
# ===============================================================================

class PatternTWDispatcher:
    """Main dispatcher that routes to appropriate TW pattern processor."""
    
    PATTERN_PROCESSORS = {
        '1plus5': OnePlusFiveProcessor,
        'comp_saod': CompSaodTWProcessor,
        'tp': TPProcessor
    }
    
    @staticmethod
    def process_sheet(df: pd.DataFrame, sheet_name: str,
                     override_enabled: bool = False,
                     override_lob: str = None,
                     override_segment: str = None,
                     override_policy_type: str = None) -> List[Dict]:
        """Main entry point for processing any TW sheet."""
        pattern = PatternTWDetector.detect_pattern(df, sheet_name)
        processor_class = PatternTWDispatcher.PATTERN_PROCESSORS.get(pattern, TPProcessor)
        
        if pattern == 'tp':
            records = processor_class.process(
                df, sheet_name, override_enabled,
                override_lob, override_segment, override_policy_type
            )
        else:
            records = processor_class.process(
                df, sheet_name, override_enabled,
                override_lob, override_segment
            )
        
        return records

# ===============================================================================
# API ENDPOINTS
# ===============================================================================

@app.get("/")
async def root():
    return {"message": "DIGIT TW Processor API", "version": "1.0"}


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel file and return available worksheets."""
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
        content = await file.read()
        xls = pd.ExcelFile(io.BytesIO(content))
        sheets = xls.sheet_names
        
        file_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        uploaded_files[file_id] = {
            "content": content,
            "filename": file.filename,
            "sheets": sheets
        }
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "sheets": sheets,
            "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


@app.post("/process")
async def process_sheet(
    file_id: str,
    sheet_name: str,
    override_enabled: bool = False,
    override_lob: Optional[str] = None,
    override_segment: Optional[str] = None,
    override_policy_type: Optional[str] = None
):
    """Process a specific worksheet and return results."""
    try:
        if file_id not in uploaded_files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
        file_data = uploaded_files[file_id]
        content = file_data["content"]
        
        if sheet_name not in file_data["sheets"]:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
        # Load sheet
        df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        # Process the sheet
        records = PatternTWDispatcher.process_sheet(
            df, sheet_name, override_enabled,
            override_lob, override_segment, override_policy_type
        )
        
        if not records:
            return {
                "success": False,
                "message": "No records extracted. Please check the sheet structure.",
                "records": [],
                "count": 0
            }
        
        # Calculate summary statistics
        states = {}
        policies = {}
        payins = []
        payouts = []
        
        for record in records:
            state = record.get("State", "Unknown")
            states[state] = states.get(state, 0) + 1
            
            policy = record.get("Policy Type", "Unknown")
            policies[policy] = policies.get(policy, 0) + 1
            
            try:
                payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
                payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
                if payin > 0:
                    payins.append(payin)
                    payouts.append(payout)
            except:
                pass
        
        avg_payin = sum(payins) / len(payins) if payins else 0
        avg_payout = sum(payouts) / len(payouts) if payouts else 0
        
        summary = {
            "total_records": len(records),
            "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
            "policies": policies,
            "average_payin": round(avg_payin, 2),
            "average_payout": round(avg_payout, 2)
        }
        
        return {
            "success": True,
            "message": f"Successfully processed {len(records)} records",
            "records": records,
            "count": len(records),
            "summary": summary
        }
        
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")


@app.post("/export")
async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
    """Export processed records to Excel file."""
    try:
        if not records:
            raise HTTPException(status_code=400, detail="No records to export")
        
        df = pd.DataFrame(records)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"TW_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, filename)
        
        df.to_excel(output_path, index=False, sheet_name='Processed')
        
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
