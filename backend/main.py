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
