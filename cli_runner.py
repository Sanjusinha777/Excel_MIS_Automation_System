"""
Command Line Runner for Automated Excel MIS Generation
Usage:
    python cli_runner.py <input_file.xlsx> [optional_output_file.xlsx]
"""

import sys
import os
from mis_engine import MISEngine
from excel_reporter import ExcelMISReporter

def run_mis_pipeline(input_path: str, output_path: str = None):
    if not os.path.exists(input_path):
        print(f"[ERROR] Input file not found: {input_path}")
        sys.exit(1)
        
    print(f"\n=======================================================")
    print(f"       EXCEL MIS REPORT AUTOMATION ENGINE")
    print(f"=======================================================")
    print(f"[*] Reading source dataset: {input_path}")
    
    engine = MISEngine(file_source=input_path)
    
    print(f"[*] Available Sheets: {', '.join(engine.sheets)}")
    print(f"[*] Selected Sheet: {engine.sheet_name}")
    print(f"[*] Records: {len(engine.cleaned_df):,} rows, {len(engine.cleaned_df.columns)} columns")
    print(f"[*] Detected Metric Columns: {engine.metric_cols}")
    print(f"[*] Detected Dimension Columns: {engine.dimension_cols}")
    print(f"[*] Detected Date Columns: {engine.date_cols}")
    print(f"[*] Primary Metric: {engine.primary_metric}")
    print(f"[*] Primary Dimension: {engine.primary_dimension}")
    print(f"[*] Primary Date: {engine.primary_date}")
    
    kpis = engine.get_summary_kpis()
    metric_name = kpis.get("metric_name", "Records")
    print(f"\n--- EXECUTIVE KPI SNAPSHOT ---")
    print(f"Total Transactions : {kpis.get('total_records', 0):,}")
    print(f"Total {metric_name}   : {kpis.get('total_value', 0):,.2f}")
    print(f"Average {metric_name} : {kpis.get('avg_value', 0):,.2f}")
    print(f"Max {metric_name}     : {kpis.get('max_value', 0):,.2f}")
    
    print(f"\n--- AUTOMATED EXECUTIVE INSIGHTS ---")
    insights = engine.generate_executive_insights()
    for ins in insights:
        clean = ins.replace("**", "")
        print(f" - {clean}")
        
    # Output path
    if not output_path:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_MIS_Report.xlsx"
        
    print(f"\n[*] Compiling multi-tab formatted Excel MIS Workbook...")
    reporter = ExcelMISReporter(engine)
    reporter.save_to_file(output_path)
    print(f"[SUCCESS] Executive MIS Report generated at: {output_path}\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python cli_runner.py <input_file.xlsx> [output_file.xlsx]")
        sys.exit(1)
        
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    run_mis_pipeline(inp, out)
