# from fastapi import FastAPI, File, UploadFile, HTTPException
# from fastapi.middleware.cors import CORSMiddleware
# from fastapi.responses import FileResponse
# import pandas as pd
# import io
# import os
# import re
# from typing import List, Dict, Tuple, Optional
# from datetime import datetime
# import traceback
# import tempfile

# app = FastAPI(title="DIGIT TW Processor API")

# # Enable CORS
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["https://digit-excel-two-wheelers.vercel.app"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

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

# # Store uploaded files temporarily
# uploaded_files = {}

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
# # API ENDPOINTS
# # ===============================================================================

# @app.get("/")
# async def root():
#     return {"message": "DIGIT TW Processor API", "version": "1.0"}


# @app.post("/upload")
# async def upload_file(file: UploadFile = File(...)):
#     """Upload an Excel file and return available worksheets."""
#     try:
#         if not file.filename.endswith(('.xlsx', '.xls')):
#             raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
#         content = await file.read()
#         xls = pd.ExcelFile(io.BytesIO(content))
#         sheets = xls.sheet_names
        
#         file_id = datetime.now().strftime("%Y%m%d_%H%M%S")
#         uploaded_files[file_id] = {
#             "content": content,
#             "filename": file.filename,
#             "sheets": sheets
#         }
        
#         return {
#             "file_id": file_id,
#             "filename": file.filename,
#             "sheets": sheets,
#             "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)."
#         }
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


# @app.post("/process")
# async def process_sheet(
#     file_id: str,
#     sheet_name: str,
#     override_enabled: bool = False,
#     override_lob: Optional[str] = None,
#     override_segment: Optional[str] = None,
#     override_policy_type: Optional[str] = None
# ):
#     """Process a specific worksheet and return results."""
#     try:
#         if file_id not in uploaded_files:
#             raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
#         file_data = uploaded_files[file_id]
#         content = file_data["content"]
        
#         if sheet_name not in file_data["sheets"]:
#             raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
#         # Load sheet
#         df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
#         # Process the sheet
#         records = PatternTWDispatcher.process_sheet(
#             df, sheet_name, override_enabled,
#             override_lob, override_segment, override_policy_type
#         )
        
#         if not records:
#             return {
#                 "success": False,
#                 "message": "No records extracted. Please check the sheet structure.",
#                 "records": [],
#                 "count": 0
#             }
        
#         # Calculate summary statistics
#         states = {}
#         policies = {}
#         payins = []
#         payouts = []
        
#         for record in records:
#             state = record.get("State", "Unknown")
#             states[state] = states.get(state, 0) + 1
            
#             policy = record.get("Policy Type", "Unknown")
#             policies[policy] = policies.get(policy, 0) + 1
            
#             try:
#                 payin = float(record.get("Payin (CD2)", "0%").replace('%', ''))
#                 payout = float(record.get("Calculated Payout", "0%").replace('%', ''))
#                 if payin > 0:
#                     payins.append(payin)
#                     payouts.append(payout)
#             except:
#                 pass
        
#         avg_payin = sum(payins) / len(payins) if payins else 0
#         avg_payout = sum(payouts) / len(payouts) if payouts else 0
        
#         summary = {
#             "total_records": len(records),
#             "states": dict(sorted(states.items(), key=lambda x: x[1], reverse=True)[:10]),
#             "policies": policies,
#             "average_payin": round(avg_payin, 2),
#             "average_payout": round(avg_payout, 2)
#         }
        
#         return {
#             "success": True,
#             "message": f"Successfully processed {len(records)} records",
#             "records": records,
#             "count": len(records),
#             "summary": summary
#         }
        
#     except Exception as e:
#         traceback.print_exc()
#         raise HTTPException(status_code=500, detail=f"Error processing sheet: {str(e)}")


# @app.post("/export")
# async def export_to_excel(file_id: str, sheet_name: str, records: List[Dict]):
#     """Export processed records to Excel file."""
#     try:
#         if not records:
#             raise HTTPException(status_code=400, detail="No records to export")
        
#         df = pd.DataFrame(records)
        
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         filename = f"TW_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
#         temp_dir = tempfile.gettempdir()
#         output_path = os.path.join(temp_dir, filename)
        
#         df.to_excel(output_path, index=False, sheet_name='Processed')
        
#         return FileResponse(
#             path=output_path,
#             filename=filename,
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
        
#     except Exception as e:
#         raise HTTPException(status_code=500, detail=f"Error exporting file: {str(e)}")


# if __name__ == "__main__":
#     import uvicorn
#     uvicorn.run(app, host="0.0.0.0", port=8000)



from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import io
import os
from typing import List, Dict, Optional
from datetime import datetime
import traceback
import tempfile
import re

app = FastAPI(title="DIGIT Multi-LOB Processor API - Enhanced TW Edition (FIXED)")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://digit-excel-two-wheelers.vercel.app"],
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
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 30%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 40%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "PO": "-3%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "PO": "88% of Payin", "REMARKS": "NIL"}
]

STATE_MAPPING = {
    "DELHI": "DELHI", "MUMBAI": "MAHARASHTRA", "PUNE": "MAHARASHTRA", "GOA": "GOA",
    "KOLKATA": "WEST BENGAL", "HYDERABAD": "TELANGANA", "AHMEDABAD": "GUJARAT",
    "TAMIL NADU": "TAMIL NADU", "TN": "TAMIL NADU", "CHENNAI": "TAMIL NADU",
    "KERALA": "KERALA", "KARNATAKA": "KARNATAKA", "BANGALORE": "KARNATAKA",
    "GUJARAT": "GUJARAT", "RAJASTHAN": "RAJASTHAN", "PUNJAB": "PUNJAB",
    "UTTAR PRADESH": "UTTAR PRADESH", "DELHI NCR": "DELHI", "REST OF INDIA": "REST OF INDIA",
    "GOOD GJ": "GUJARAT", "BAD GJ": "GUJARAT", "ROM1": "REST OF MAHARASHTRA",
    "ROM2": "REST OF MAHARASHTRA", "GOOD TN": "TAMIL NADU", "GOOD MP": "MADHYA PRADESH",
    "ANDAMAN": "ANDAMAN AND NICOBAR ISLANDS", "BIHAR": "BIHAR", "CHANDIGARH": "CHANDIGARH",
    "HIMACHAL": "HIMACHAL PRADESH", "J&K": "JAMMU AND KASHMIR", "JHARKHAND": "JHARKHAND",
    "MAHARASHTRA": "MAHARASHTRA", "HARYANA": "HARYANA", "WEST BENGAL": "WEST BENGAL",
    "UP": "UTTAR PRADESH", "UK": "UTTARAKHAND", "BR": "BIHAR", "DL": "DELHI", 
    "GJ": "GUJARAT", "MH": "MAHARASHTRA", "WB": "WEST BENGAL", "NCR": "DELHI NCR",
    "ROM": "REST OF MAHARASHTRA", "RJ": "RAJASTHAN", "OR": "ODISHA", "AP": "ANDHRA PRADESH",
    "ASSAM": "ASSAM", "KA": "KARNATAKA", "JH": "JHARKHAND", "NE": "NORTH EAST"
}

uploaded_files = {}

# ===============================================================================
# CORE CALCULATION FUNCTIONS
# ===============================================================================

def get_payin_category(payin: float):
    if payin <= 20: return "Payin Below 20%"
    elif payin <= 30: return "Payin 21% to 30%"
    elif payin <= 50: return "Payin 31% to 50%"
    else: return "Payin Above 50%"

def safe_float(value):
    if pd.isna(value): return None
    val_str = str(value).strip().upper()
    if val_str in ["D", "NA", "", "NAN", "NONE", "DECLINE"]: return None
    try:
        # Handle both "20%" string and 0.20 float
        if isinstance(value, str):
            val_str = value.strip().replace('%', '')
        num = float(val_str)
        if num < 0:
            return None  # Skip negative values
        # Convert decimals (0.20) to percentages (20)
        if 0 < num <= 1:
            return round(num * 100, 2)
        elif num > 1:
            return round(num, 2)
        return 0  # Return 0 for exact 0
    except:
        return None

def extract_state(cluster_name: str) -> str:
    if pd.isna(cluster_name):
        return "UNKNOWN"
    cluster_upper = str(cluster_name).upper().strip()
    for key, val in STATE_MAPPING.items():
        if key in cluster_upper:
            return val
    return "REST OF INDIA"

def get_formula_from_data(lob: str, segment: str, policy_type: str, payin: float):
    segment_key = segment.upper()
    if lob == "TW":
        if segment_key == "1+5":
            # 1+5 always uses 90% of Payin
            return "90% of Payin", round(payin * 0.9, 2)
        elif "TP" in policy_type.upper():
            segment_key = "TW TP"
        else:
            segment_key = "TW SAOD + COMP"
    elif lob == "PVT CAR":
        segment_key = "PVT CAR TP" if policy_type == "TP" else "PVT CAR COMP + SAOD"
    elif lob in ["TAXI", "CV", "BUS", "MISD"]:
        segment_key = segment.upper()

    payin_category = get_payin_category(payin)
    matching_rule = None
    for rule in FORMULA_DATA:
        if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
            if rule["REMARKS"] == payin_category or rule["REMARKS"] == "NIL":
                matching_rule = rule
                break

    if not matching_rule and payin > 20:
        for rule in FORMULA_DATA:
            if rule["LOB"] == lob and rule["SEGMENT"] == segment_key:
                if (rule["REMARKS"] == "Payin Above 20%" or
                    (payin > 30 and rule["REMARKS"] == "Payin Above 30%") or
                    (payin > 40 and rule["REMARKS"] == "Payin Above 40%") or
                    (payin > 50 and rule["REMARKS"] == "Payin Above 50%")):
                    matching_rule = rule
                    break

    if not matching_rule:
        deduction = 2 if payin <= 20 else 3 if payin <= 30 else 4 if payin <= 50 else 5
        return f"-{deduction}%", round(payin - deduction, 2)

    formula = matching_rule["PO"]
    if "% of Payin" in formula:
        perc_str = formula.split("%")[0].replace("Less ", "").strip()
        percentage = float(perc_str)
        if "Less" in formula:
            payout = round(payin - percentage, 2)
        else:
            payout = round(payin * percentage / 100, 2)
    elif formula.startswith("-"):
        deduction = float(formula.replace("%", "").replace("-", ""))
        payout = round(payin - deduction, 2)
    else:
        payout = round(payin - 2, 2)

    return formula, payout

def calculate_payout_with_formula(lob: str, segment: str, policy_type: str, payin: float):
    if payin == 0:
        return 0, "0% (No Payin)", "Payin is 0"
    formula, payout = get_formula_from_data(lob, segment, policy_type, payin)
    return payout, formula, f"Match: LOB={lob}, Segment={segment}, Policy={policy_type}, {get_payin_category(payin)}"

def clean_year_label(header: str) -> str:
    """Convert '1 year CD2' or '2 year CD2' into 'Year 1'"""
    match = re.search(r'(\d+)\s*year', header, re.IGNORECASE)
    if match:
        year_num = match.group(1)
        return f"Year {year_num}"
    return header.strip()

# ===============================================================================
# ENHANCED TW SHEET PROCESSORS (ALL PATTERNS) - FIXED
# ===============================================================================

def detect_tw_pattern(df):
    """
    Auto-detect which TW pattern this sheet follows.
    FIXED: Proper type conversion to avoid string concatenation errors.
    """
    # Check first 10 rows for clues - FIXED TYPE CONVERSION
    sample_rows = []
    for i in range(min(10, len(df))):
        row_values = []
        for val in df.iloc[i]:
            if pd.notna(val):
                row_values.append(str(val).upper())
        sample_rows.append(" ".join(row_values))
    
    sample_text = " ".join(sample_rows)
    
    # Check for multi-year CD2 columns
    if "1 YEAR CD2" in sample_text and "2 YEAR CD2" in sample_text:
        return "tw_saod_multiyear_cd2"
    
    # Check for CD2 columns (flexible SAOD)
    if "CD2" in sample_text and "SEGMENT" in sample_text and "CLUSTER" in sample_text:
        # Could be flexible SAOD or 1+5
        if "MAKE" in sample_text:
            return "tw_1plus5_with_make"
        elif any(x in sample_text for x in ["1+5", "1PLUS5"]):
            return "tw_1plus5_simple"
        else:
            return "tw_saod_flexible_cd2"
    
    # Check for Make column (1+5 with Make)
    if "MAKE" in sample_text and "AGENCY/PB CLUSTERS" in sample_text:
        return "tw_1plus5_with_make"
    
    # Check for simple 1+5 pattern
    if ("AGENCY/PB CLUSTERS" in sample_text or "CLUSTER" in sample_text) and \
       ("1+5" in sample_text or "CD2" in sample_text):
        return "tw_1plus5_simple"
    
    # Default to TW simple if nothing else matches
    return "tw_simple"

def process_tw_1plus5_with_make(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Pattern: July, August, Sept, Oct - has Make, Seg, CD2, Formula Type columns"""
    records = []
    
    # Find header row
    header_row = None
    for i in range(min(10, df.shape[0])):
        row_values = [str(val).upper() if pd.notna(val) else "" for val in df.iloc[i]]
        if "AGENCY/PB CLUSTERS" in " ".join(row_values) or "MAKE" in " ".join(row_values):
            header_row = i
            break
    
    if header_row is None:
        header_row = 0
    
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
    
    cluster_col = df.columns[0]
    make_col = df.columns[1] if len(df.columns) > 1 else None
    segment_col = df.columns[2] if len(df.columns) > 2 else None
    cd2_col = df.columns[4] if len(df.columns) > 4 else None
    formula_type_col = df.columns[5] if len(df.columns) > 5 else None
    
    lob_final = override_lob if override_enabled and override_lob else "TW"
    segment_final = override_segment if override_enabled and override_segment else "1+5"
    
    for idx, row in df.iterrows():
        cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
        if not cluster or cluster.upper() in ["", "TOTAL", "GRAND TOTAL", "AGENCY/PB CLUSTERS"]:
            continue
        
        payin = safe_float(row[cd2_col]) if cd2_col else None
        if payin is None or payin <= 0:
            continue
        
        state = extract_state(cluster)
        payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", payin)
        
        make = str(row[make_col]).strip() if make_col and pd.notna(row[make_col]) else ""
        formula_type = str(row[formula_type_col]).strip() if formula_type_col and pd.notna(row[formula_type_col]) else ""
        combined_remarks = " | ".join(filter(None, [make, formula_type]))
        
        records.append({
            "State": state,
            "Location/Cluster": cluster,
            "Make": make,
            "Original Segment": str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else "",
            "Mapped Segment": segment_final,
            "LOB": lob_final,
            "Policy Type": "TP",
            "Status": "STP",
            "Payin (CD2)": f"{payin:.2f}%",
            "Calculated Payout": f"{payout:.2f}%",
            "Formula Used": formula,
            "Payin Category": get_payin_category(payin),
            "Remarks": combined_remarks
        })
    
    return records

def process_tw_1plus5_simple(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Pattern: Jan, Feb, March, May - simple 1+5 format"""
    records = []
    
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = [str(col).strip() if pd.notna(col) else f"Col_{i}" for i, col in enumerate(df.columns)]
    
    # Find CD2 column
    cd2_col = None
    for col in df.columns:
        col_upper = str(col).upper()
        if "CD2" in col_upper or "1+5" in col_upper:
            cd2_col = col
            break
    
    if cd2_col is None:
        cd2_col = df.columns[3] if len(df.columns) > 3 else df.columns[-1]
    
    cluster_col = df.columns[0]
    segment_col = df.columns[1] if len(df.columns) > 1 else None
    
    lob_final = override_lob if override_enabled and override_lob else "TW"
    segment_final = override_segment if override_enabled and override_segment else "1+5"
    
    for idx, row in df.iterrows():
        cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
        if not cluster or cluster.upper() in ["AGENCY/PB CLUSTERS", "TOTAL", "GRAND TOTAL"]:
            continue
        
        payin = safe_float(row[cd2_col])
        if payin is None:
            continue
        
        state = extract_state(cluster)
        payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", payin)
        
        records.append({
            "State": state,
            "Location/Cluster": cluster,
            "Original Segment": str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else "",
            "Mapped Segment": segment_final,
            "LOB": lob_final,
            "Policy Type": "TP",
            "Status": "STP",
            "Payin (CD2)": f"{payin:.2f}%",
            "Payin Category": get_payin_category(payin),
            "Calculated Payout": f"{payout:.2f}%",
            "Formula Used": formula
        })
    
    return records

def process_tw_saod_multiyear_cd2(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Pattern: SAOD with multiple year CD2 columns (June, July, Aug, Sept, Oct)"""
    records = []
    
    # Find header row
    header_row = None
    for i in range(min(10, df.shape[0])):
        row_str = " ".join([str(val).upper() if pd.notna(val) else "" for val in df.iloc[i]])
        if "CLUSTER" in row_str and "1 YEAR CD2" in row_str:
            header_row = i
            break
    
    if header_row is None:
        header_row = 0
    
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
    
    cluster_col = df.columns[0]
    segment_col = df.columns[1] if len(df.columns) > 1 else None
    cd2_year1_col = df.columns[5] if len(df.columns) > 5 else None
    cd2_year2_col = df.columns[6] if len(df.columns) > 6 else None
    cd2_year3_col = df.columns[7] if len(df.columns) > 7 else None
    cd2_year4_col = df.columns[8] if len(df.columns) > 8 else None
    
    lob_final = override_lob if override_enabled and override_lob else "TW"
    segment_final = override_segment if override_enabled and override_segment else "TW SAOD + COMP"
    
    year_columns = [
        (1, cd2_year1_col, "1 year CD2"),
        (2, cd2_year2_col, "2 year CD2"),
        (3, cd2_year3_col, "3rd year CD2"),
        (4, cd2_year4_col, "4th year CD2")
    ]
    
    for idx, row in df.iterrows():
        cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
        if not cluster or cluster.upper() in ["CLUSTER", "TOTAL", "GRAND TOTAL"]:
            continue
        
        original_segment = str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else ""
        state = extract_state(cluster)
        
        for year_num, cd2_col, remark_text in year_columns:
            if cd2_col is None:
                continue
            payin = safe_float(row[cd2_col])
            if payin is None or payin == 0:
                continue
            
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "SAOD", payin)
            
            records.append({
                "State": state,
                "Location/Cluster": cluster,
                "Original Segment": original_segment,
                "Mapped Segment": segment_final,
                "LOB": lob_final,
                "Policy Type": "SAOD",
                "Status": "STP",
                "Year": f"Year {year_num}",
                "Payin (CD2)": f"{payin:.2f}%",
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Payin Category": get_payin_category(payin),
                "Remarks": remark_text
            })
    
    return records

def process_tw_saod_flexible_cd2(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Pattern: SAOD with flexible number of CD2 columns (June Pattern 2)"""
    records = []
    
    # Find header row
    header_row = None
    for i in range(min(10, df.shape[0])):
        row_str = " ".join([str(val).upper() if pd.notna(val) else "" for val in df.iloc[i]])
        if "CLUSTER" in row_str and "SEGMENT" in row_str:
            header_row = i
            break
    
    if header_row is None:
        header_row = 0
    
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(col).strip() if pd.notna(col) else f"Unnamed_{i}" for i, col in enumerate(df.columns)]
    
    cluster_col = df.columns[0]
    segment_col = df.columns[1] if len(df.columns) > 1 else None
    
    # Find all CD2 columns
    cd2_columns = []
    for col in df.columns:
        if "CD2" in str(col).upper():
            clean_label = clean_year_label(str(col))
            cd2_columns.append((col, clean_label, str(col)))
    
    if not cd2_columns:
        return []
    
    lob_final = override_lob if override_enabled and override_lob else "TW"
    segment_final = override_segment if override_enabled and override_segment else "TW SAOD + COMP"
    
    for idx, row in df.iterrows():
        cluster = str(row[cluster_col]).strip() if pd.notna(row[cluster_col]) else ""
        if not cluster or cluster.upper() in ["CLUSTER", "TOTAL", "GRAND TOTAL"]:
            continue
        
        original_segment = str(row[segment_col]).strip() if segment_col and pd.notna(row[segment_col]) else ""
        state = extract_state(cluster)
        
        for col_name, year_label, original_header in cd2_columns:
            payin = safe_float(row[col_name])
            if payin is None or payin == 0:
                continue
            
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "SAOD", payin)
            
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
                "Calculated Payout": f"{payout:.2f}%",
                "Formula Used": formula,
                "Payin Category": get_payin_category(payin),
                "Remarks": original_header
            })
    
    return records

def process_tw_sheet_unified(df, override_enabled, override_lob, override_segment, override_policy_type):
    """
    Unified TW processor that auto-detects pattern and routes to appropriate handler
    """
    pattern = detect_tw_pattern(df)
    
    print(f"🔍 Detected TW Pattern: {pattern}")
    
    if pattern == "tw_1plus5_with_make":
        return process_tw_1plus5_with_make(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == "tw_1plus5_simple":
        return process_tw_1plus5_simple(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == "tw_saod_multiyear_cd2":
        return process_tw_saod_multiyear_cd2(df, override_enabled, override_lob, override_segment, override_policy_type)
    elif pattern == "tw_saod_flexible_cd2":
        return process_tw_saod_flexible_cd2(df, override_enabled, override_lob, override_segment, override_policy_type)
    else:
        # Fallback to simple processor
        return process_tw_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)

# ===============================================================================
# ORIGINAL SHEET PROCESSORS (kept as fallback)
# ===============================================================================

def process_tw_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Fallback TW processor"""
    records = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
        cluster = str(row.iloc[0]).strip()
        segmentation = str(row.iloc[1]).strip() if len(row) > 1 else ""
        comp_cd2 = safe_float(row.iloc[3]) if len(row) > 3 else None
        satp_cd2 = safe_float(row.iloc[4]) if len(row) > 4 else None

        state = extract_state(cluster)
        lob_final = override_lob if override_enabled and override_lob else "TW"
        segment_final = override_segment if override_enabled and override_segment else "TW"

        if comp_cd2 is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, comp_cd2)
            records.append({"State": state, "Location/Cluster": cluster, "Original Segment": f"TW {segmentation}",
                            "Mapped Segment": segment_final, "LOB": lob_final, "Policy Type": policy_type,
                            "Payin (CD2)": f"{comp_cd2:.2f}%", "Payin Category": get_payin_category(comp_cd2),
                            "Calculated Payout": f"{payout:.2f}%", "Formula Used": formula, "Rule Explanation": rule_exp})

        if satp_cd2 is not None:
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", satp_cd2)
            records.append({"State": state, "Location/Cluster": cluster, "Original Segment": f"TW {segmentation}",
                            "Mapped Segment": segment_final, "LOB": lob_final, "Policy Type": "TP",
                            "Payin (CD2)": f"{satp_cd2:.2f}%", "Payin Category": get_payin_category(satp_cd2),
                            "Calculated Payout": f"{payout:.2f}%", "Formula Used": formula, "Rule Explanation": rule_exp})
    return records

def process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    records = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]): continue
        city = str(row.iloc[0]).strip()
        fuel = str(row.iloc[2]).strip() if len(row) > 2 else "Electric"
        cvod_cd2 = safe_float(row.iloc[6]) if len(row) > 6 else None
        cvtp_cd2 = safe_float(row.iloc[7]) if len(row) > 7 else None

        state = next((v for k, v in STATE_MAPPING.items() if k.upper() in city.upper()), "UNKNOWN")
        lob_final = override_lob if override_enabled and override_lob else "TAXI"
        segment_final = override_segment if override_enabled and override_segment else "TAXI"

        if cvod_cd2 is not None:
            policy_type = override_policy_type or "Comp"
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, cvod_cd2)
            records.append({"State": state, "Location/Cluster": city, "Original Segment": f"Taxi {fuel}",
                            "Mapped Segment": segment_final, "LOB": lob_final, "Policy Type": policy_type,
                            "Payin (CD2)": f"{cvod_cd2:.2f}%", "Payin Category": get_payin_category(cvod_cd2),
                            "Calculated Payout": f"{payout:.2f}%", "Formula Used": formula, "Rule Explanation": rule_exp})

        if cvtp_cd2 is not None:
            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, "TP", cvtp_cd2)
            records.append({"State": state, "Location/Cluster": city, "Original Segment": f"Taxi {fuel}",
                            "Mapped Segment": segment_final, "LOB": lob_final, "Policy Type": "TP",
                            "Payin (CD2)": f"{cvtp_cd2:.2f}%", "Payin Category": get_payin_category(cvtp_cd2),
                            "Calculated Payout": f"{payout:.2f}%", "Formula Used": formula, "Rule Explanation": rule_exp})
    return records

def process_4w_satp_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    records = []
    for _, row in df.iterrows():
        if pd.isna(row.get('Cluster')): continue
        cluster = str(row['Cluster']).strip()
        payin = safe_float(row.get('CD2'))
        if payin is None: continue

        state = next((v for k, v in STATE_MAPPING.items() if k.upper() in cluster.upper()), "UNKNOWN")
        lob_final = override_lob if override_enabled and override_lob else "PVT CAR"
        segment_final = override_segment if override_enabled and override_segment else "PVT CAR TP"
        policy_type = override_policy_type or "TP"

        payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type, payin)
        records.append({"State": state, "Location/Cluster": cluster, "Original Segment": "PVT CAR TP",
                        "Mapped Segment": segment_final, "LOB": lob_final, "Policy Type": policy_type,
                        "Payin (CD2)": f"{payin:.2f}%", "Payin Category": get_payin_category(payin),
                        "Calculated Payout": f"{payout:.2f}%", "Formula Used": formula, "Rule Explanation": rule_exp})
    return records

def process_school_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Enhanced school bus processor"""
    records = []
    
    table_start_row = None
    for i in range(min(30, len(df))):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower() if pd.notna(df.iloc[i, j]) else ""
            if "school bus" in cell:
                table_start_row = i
                break
        if table_start_row is not None: break
    
    if table_start_row is None: 
        return []

    seating_row = None
    seating_col = -1
    for i in range(table_start_row + 1, min(table_start_row + 5, len(df))):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower() if pd.notna(df.iloc[i, j]) else ""
            if "seating capacity" in cell:
                seating_row = i
                seating_col = j
                break
        if seating_row is not None: break

    data_start_row = (seating_row + 1) if seating_row is not None else (table_start_row + 2)

    if seating_col > 0:
        state_col = 0
        rto_col = 1
        payin_start = seating_col + 1
        contract_types = [str(df.iloc[seating_row, k]).strip() for k in range(payin_start, len(df.columns)) 
                         if str(df.iloc[seating_row, k]).strip()]
        contracts = [(payin_start + idx, ct) for idx, ct in enumerate(contract_types)]
    else:
        state_col = 0
        rto_col = 1
        contracts = [(2, "In name of School"), (3, "On Contract (Transporter)"), 
                    (4, "On Contract (Individual)"), (5, "Contract transporter")]

    current_state = ""
    for row_idx in range(data_start_row, len(df)):
        first_cell = str(df.iloc[row_idx, state_col]).strip().lower() if pd.notna(df.iloc[row_idx, state_col]) else ""
        if any(kw in first_cell for kw in ["staff bus", "note"]): 
            break

        state_val = str(df.iloc[row_idx, state_col]).strip() if pd.notna(df.iloc[row_idx, state_col]) else ""
        if state_val: 
            current_state = state_val

        rto_cluster = str(df.iloc[row_idx, rto_col]).strip() if pd.notna(df.iloc[row_idx, rto_col]) else ""
        if not rto_cluster: 
            continue

        for col_idx, contract_type in contracts:
            if col_idx >= len(df.columns):
                continue
            payin = safe_float(df.iloc[row_idx, col_idx])
            if payin is None: 
                continue

            state_mapped = next((v for k, v in STATE_MAPPING.items() if k.upper() in current_state.upper()), 
                              current_state.upper())
            lob_final = override_lob if override_enabled and override_lob else "BUS"
            segment_final = override_segment if override_enabled and override_segment else "SCHOOL BUS"
            policy_type_final = override_policy_type or "Comp"

            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)
            record = {
                "State": state_mapped.upper(), 
                "Location/Cluster": f"{current_state} - {rto_cluster}",
                "Original Segment": f"School Bus - {contract_type}", 
                "Mapped Segment": segment_final,
                "LOB": lob_final, 
                "Policy Type": policy_type_final,
                "Payin (CD2)": f"{payin:.2f}%", 
                "Payin Category": get_payin_category(payin),
                "Calculated Payout": f"{payout:.2f}%", 
                "Formula Used": formula, 
                "Rule Explanation": rule_exp
            }
            
            if seating_col > 0 and seating_col < len(df.columns):
                seating = str(df.iloc[row_idx, seating_col]).strip()
                if seating: 
                    record["Original Segment"] += f" ({seating})"
            
            records.append(record)
    
    return records

def process_staff_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Enhanced staff bus processor"""
    records = []
    
    table_start_row = None
    for i in range(min(30, len(df))):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).strip().lower() if pd.notna(df.iloc[i, j]) else ""
            if "staff bus" in cell:
                table_start_row = i
                break
        if table_start_row is not None: break
    
    if table_start_row is None: 
        return []

    data_start_row = table_start_row + 2
    contracts = [(1, "In name of Company"), (2, "Contract (Transport)"), (3, "Contract (Individual)")]

    for row_idx in range(data_start_row, len(df)):
        rto_val = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        if not rto_val or any(kw in rto_val.lower() for kw in ["note", "permit", "validation", "exception", "above grid"]): 
            continue

        for col_idx, contract_type in contracts:
            if col_idx >= len(df.columns):
                continue
                
            cell_value = str(df.iloc[row_idx, col_idx]).strip()
            if not cell_value or "decline" in cell_value.lower(): 
                continue

            payin = None
            if "CD2" in cell_value.upper():
                for part in cell_value.split("/"):
                    if "CD2" in part.upper():
                        cd2_cleaned = part.replace("CD2", "").replace("cd2", "").strip()
                        payin = safe_float(cd2_cleaned)
                        break
            
            if payin is None: 
                continue

            state_mapped = next((v for k, v in STATE_MAPPING.items() if k.upper() in rto_val.upper()), "UNKNOWN")
            lob_final = override_lob if override_enabled and override_lob else "BUS"
            segment_final = override_segment if override_enabled and override_segment else "STAFF BUS"
            policy_type_final = override_policy_type or "Comp"

            payout, formula, rule_exp = calculate_payout_with_formula(lob_final, segment_final, policy_type_final, payin)
            records.append({
                "State": state_mapped.upper(), 
                "Location/Cluster": rto_val,
                "Original Segment": f"Staff Bus - {contract_type}", 
                "Mapped Segment": segment_final,
                "LOB": lob_final, 
                "Policy Type": policy_type_final,
                "Payin (CD2)": f"{payin:.2f}%", 
                "Payin Category": get_payin_category(payin),
                "Calculated Payout": f"{payout:.2f}%", 
                "Formula Used": formula, 
                "Rule Explanation": rule_exp
            })
    
    return records

def process_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type):
    """Combined bus processor"""
    school_records = process_school_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    staff_records = process_staff_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
    return school_records + staff_records

# ===============================================================================
# SHEET TYPE DETECTION
# ===============================================================================

def detect_sheet_type(sheet_name: str) -> str:
    """Detect the type of sheet based on its name."""
    name_lower = sheet_name.lower()
    
    # TW detection (most specific first)
    if any(keyword in name_lower for keyword in ['tw', 'two wheeler', '2w', '2 wheeler', '1+5', 'saod']):
        return "tw"
    
    # Bus detection
    if any(keyword in name_lower for keyword in ['bus', 'school', 'staff']):
        return "bus"
    
    # Taxi/Electric detection
    if any(keyword in name_lower for keyword in ['taxi', 'electric', 'ev', 'e-rickshaw']):
        return "taxi"
    
    # 4W SATP detection
    if any(keyword in name_lower for keyword in ['4w', 'satp', 'pvt car tp', 'four wheeler']):
        return "4w_satp"
    
    return "unknown"

def get_sheet_preview(df, sheet_type: str, max_rows: int = 5):
    """Generate a preview of the sheet data."""
    preview = {
        "columns": [],
        "sample_data": [],
        "total_rows": len(df)
    }
    
    if sheet_type == "4w_satp":
        preview["columns"] = df.columns.tolist()[:6]
    else:
        preview["columns"] = [f"Column {i+1}" for i in range(min(6, df.shape[1]))]
    
    sample_df = df.head(max_rows)
    for _, row in sample_df.iterrows():
        row_data = [str(val)[:30] if pd.notna(val) else "" for val in row.iloc[:6]]
        preview["sample_data"].append(row_data)
    
    return preview

# ===============================================================================
# API ENDPOINTS
# ===============================================================================

@app.get("/")
async def root():
    return {
        "message": "DIGIT Multi-LOB Processor API - Enhanced TW Edition (FIXED)",
        "version": "3.0.1",
        "status": "Fixed type conversion error in detect_tw_pattern",
        "features": [
            "Auto-detects ALL TW patterns (1+5 and SAOD variations)",
            "Handles monthly format changes automatically",
            "Supports Multi-Year CD2 columns",
            "Enhanced Bus processors",
            "Fixed: Type conversion error resolved"
        ],
        "endpoints": [
            "/upload - Upload Excel file and get worksheet list",
            "/process - Process a selected worksheet",
            "/export - Export processed data to Excel"
        ]
    }

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload file and return available worksheets with previews."""
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
        
        content = await file.read()
        xls = pd.ExcelFile(io.BytesIO(content))
        sheets = xls.sheet_names
        
        sheet_info = []
        for sheet in sheets:
            sheet_type = detect_sheet_type(sheet)
            
            if sheet_type == "4w_satp":
                df_preview = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=0, nrows=5)
            else:
                df_preview = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None, nrows=5)
            
            preview = get_sheet_preview(df_preview, sheet_type)
            
            sheet_info.append({
                "name": sheet,
                "type": sheet_type,
                "type_display": {
                    "bus": "Bus (School + Staff)",
                    "tw": "Two Wheeler (Auto-Pattern)",
                    "taxi": "Taxi / Electric",
                    "4w_satp": "4W SATP",
                    "unknown": "Unknown Type"
                }.get(sheet_type, "Unknown Type"),
                "icon": {
                    "bus": "🚌",
                    "tw": "🏍️",
                    "taxi": "🚕",
                    "4w_satp": "🚗",
                    "unknown": "📄"
                }.get(sheet_type, "📄"),
                "preview": preview
            })
        
        file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        uploaded_files[file_id] = {
            "content": content,
            "filename": file.filename,
            "sheets": sheets,
            "sheet_info": sheet_info
        }
        
        auto_selected = len(sheets) == 1
        auto_selected_sheet = sheets[0] if auto_selected else None
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "total_sheets": len(sheets),
            "sheet_info": sheet_info,
            "auto_selected": auto_selected,
            "auto_selected_sheet": auto_selected_sheet,
            "message": f"File uploaded successfully. Found {len(sheets)} worksheet(s)." + 
                      (f" Auto-selected '{auto_selected_sheet}' for processing." if auto_selected else " Please select a worksheet to process.")
        }
        
    except Exception as e:
        traceback.print_exc()
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
    """Process a specific worksheet with auto-pattern detection."""
    try:
        if file_id not in uploaded_files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file again.")
        
        file_data = uploaded_files[file_id]
        content = file_data["content"]
        
        if sheet_name not in file_data["sheets"]:
            raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in file")
        
        sheet_type = detect_sheet_type(sheet_name)
        
        if sheet_type == "4w_satp":
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=0)
        else:
            df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        # Route to appropriate processor
        if sheet_type == "bus":
            records = process_bus_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
            processor_name = "Bus (School + Staff)"
        elif sheet_type == "tw":
            # Use unified TW processor with auto-pattern detection
            records = process_tw_sheet_unified(df, override_enabled, override_lob, override_segment, override_policy_type)
            pattern = detect_tw_pattern(df)
            processor_name = f"Two Wheeler ({pattern.replace('_', ' ').title()})"
        elif sheet_type == "taxi":
            records = process_electric_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
            processor_name = "Taxi / Electric"
        elif sheet_type == "4w_satp":
            records = process_4w_satp_sheet(df, override_enabled, override_lob, override_segment, override_policy_type)
            processor_name = "4W SATP"
        else:
            return {
                "success": False,
                "message": f"Unable to auto-detect processor for sheet '{sheet_name}'. Please use manual override settings.",
                "records": [],
                "count": 0,
                "sheet_type": sheet_type
            }
        
        if not records:
            return {
                "success": False,
                "message": "No records extracted. Please check the sheet structure or try with override settings.",
                "records": [],
                "count": 0,
                "processor": processor_name
            }
        
        # Calculate summary statistics
        states = {}
        lobs = {}
        policies = {}
        payins = []
        payouts = []
        
        for record in records:
            state = record.get("State", "Unknown")
            states[state] = states.get(state, 0) + 1
            
            lob = record.get("LOB", "Unknown")
            lobs[lob] = lobs.get(lob, 0) + 1
            
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
            "lobs": lobs,
            "policies": policies,
            "average_payin": round(avg_payin, 2),
            "average_payout": round(avg_payout, 2),
            "processor": processor_name,
            "sheet_type": sheet_type,
            "sheet_name": sheet_name
        }
        
        return {
            "success": True,
            "message": f"Successfully processed {len(records)} records from '{sheet_name}' using {processor_name} processor",
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
        filename = f"MultiLOB_Processed_{sheet_name.replace(' ', '_')}_{timestamp}.xlsx"
        
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
    print("\n" + "="*80)
    print("DIGIT Multi-LOB Processor API - Enhanced TW Edition (FIXED)")
    print("="*80)
    print("Version: 3.0.1")
    print("Status: Type conversion error FIXED")
    print("="*80 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)


