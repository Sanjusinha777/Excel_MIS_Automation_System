"""
Core MIS (Management Information System) Analysis Engine
Automatically inspects, cleans, and generates analytical summaries from any Excel/CSV dataset.
"""

import os
import io
import re
import numpy as np
import pandas as pd
from typing import Dict, List, Any, Optional, Tuple

class MISEngine:
    def __init__(self, file_source=None, sheet_name=None):
        self.file_source = file_source
        self.sheet_name = sheet_name
        self.raw_df: Optional[pd.DataFrame] = None
        self.cleaned_df: Optional[pd.DataFrame] = None
        self.sheets: List[str] = []
        
        # Classification of columns
        self.date_cols: List[str] = []
        self.metric_cols: List[str] = []
        self.dimension_cols: List[str] = []
        self.id_cols: List[str] = []
        
        # Auto-detected defaults
        self.primary_metric: Optional[str] = None
        self.primary_dimension: Optional[str] = None
        self.primary_date: Optional[str] = None

        if file_source is not None:
            self.load_data(file_source, sheet_name)

    def load_data(self, file_source, sheet_name=None) -> pd.DataFrame:
        """Load dataset from file path, buffer, or file-like object with multi-encoding and format fallback."""
        import io
        self.file_source = file_source
        
        # Read raw bytes if it's a file-like object so we can retry multiple loaders safely
        raw_bytes = None
        file_name = getattr(file_source, 'name', str(file_source))
        
        if hasattr(file_source, 'read'):
            if hasattr(file_source, 'seek'):
                file_source.seek(0)
            raw_bytes = file_source.read()
            if hasattr(file_source, 'seek'):
                file_source.seek(0)
        elif isinstance(file_source, str) and os.path.exists(file_source):
            with open(file_source, 'rb') as f:
                raw_bytes = f.read()

        # Determine priority by file extension if available
        ext = os.path.splitext(file_name)[1].lower() if file_name else ""
        is_csv = ext in ['.csv', '.tsv', '.txt']
        is_excel = ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']

        loaded = False
        last_error = None

        def try_excel():
            nonlocal loaded, last_error
            excel_engines = ['openpyxl', 'xlrd', None]
            for engine_name in excel_engines:
                try:
                    src = io.BytesIO(raw_bytes) if raw_bytes is not None else file_source
                    excel_file = pd.ExcelFile(src, engine=engine_name) if engine_name else pd.ExcelFile(src)
                    self.sheets = excel_file.sheet_names
                    if sheet_name is None or sheet_name not in self.sheets:
                        self.sheet_name = self.sheets[0]
                    else:
                        self.sheet_name = sheet_name
                    self.raw_df = pd.read_excel(excel_file, sheet_name=self.sheet_name)
                    loaded = True
                    return True
                except Exception as e:
                    last_error = e
            return False

        def try_csv():
            nonlocal loaded, last_error
            encodings_to_try = ['utf-8', 'cp1252', 'latin1', 'utf-8-sig', 'iso-8859-1', 'gbk', 'windows-1250']
            for enc in encodings_to_try:
                try:
                    src = io.BytesIO(raw_bytes) if raw_bytes is not None else file_source
                    self.raw_df = pd.read_csv(src, encoding=enc, engine='python', on_bad_lines='skip')
                    self.sheets = ["CSV_Data"]
                    self.sheet_name = "CSV_Data"
                    loaded = True
                    return True
                except Exception as e:
                    last_error = e
            # Ultimate fallback with encoding_errors='replace'
            try:
                src = io.BytesIO(raw_bytes) if raw_bytes is not None else file_source
                self.raw_df = pd.read_csv(src, encoding='utf-8', encoding_errors='replace', engine='python', on_bad_lines='skip')
                self.sheets = ["CSV_Data"]
                self.sheet_name = "CSV_Data"
                loaded = True
                return True
            except Exception as e:
                last_error = e
            return False

        if is_csv:
            if not try_csv():
                try_excel()
        elif is_excel:
            if not try_excel():
                try_csv()
        else:
            # Unknown extension, try excel then csv
            if not try_excel():
                try_csv()

        if not loaded or self.raw_df is None:
            raise ValueError(f"Unable to read file: {str(last_error)}")

        self.clean_and_classify()
        return self.cleaned_df

    def clean_and_classify(self):
        """Clean headers, cast datatypes where obvious, and classify column roles."""
        df = self.raw_df.copy()
        
        # 1. Clean column names (strip whitespace)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 2. Check and parse potential dates
        self.date_cols = []
        date_keywords = r'(date|time|period|month|quarter|year|day|timestamp|created|modified|updated|doj|dob)'
        
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                self.date_cols.append(col)
                continue
            
            # If string or object, check if column name hints at date or values look like dates
            if df[col].dtype == 'object' or pd.api.types.is_string_dtype(df[col]):
                non_null = df[col].dropna()
                if len(non_null) > 0:
                    name_match = re.search(date_keywords, col, re.IGNORECASE)
                    sample = non_null.head(10)
                    try:
                        import warnings
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore")
                            parsed = pd.to_datetime(sample, errors='coerce', format='mixed')
                            if (name_match and parsed.notna().sum() >= len(sample) * 0.7) or (parsed.notna().sum() == len(sample) and sample.astype(str).str.contains(r'[-/:]').all()):
                                df[col] = pd.to_datetime(df[col], errors='coerce', format='mixed')
                                self.date_cols.append(col)
                    except Exception:
                        pass

        # 3. Classify Numeric Metrics, IDs, and Categorical Dimensions
        self.metric_cols = []
        self.dimension_cols = []
        self.id_cols = []
        
        id_pattern = r'(id|code|number|num|no|key|sku|ref|phone|mobile|zip|pin|account|invoice|bill|challan)'
        
        for col in df.columns:
            if col in self.date_cols:
                continue
                
            is_numeric = pd.api.types.is_numeric_dtype(df[col])
            unique_count = df[col].nunique(dropna=True)
            total_rows = len(df)
            col_lower = col.lower()
            
            # Check for ID columns
            is_id_name = bool(re.search(id_pattern, col_lower))
            if is_id_name and (unique_count > total_rows * 0.6 or not is_numeric):
                self.id_cols.append(col)
                continue
            
            if is_numeric:
                # If numeric column has very few unique values (< 4) and total_rows > 20, might be a category/flag
                if unique_count <= 4 and total_rows > 20 and not any(k in col_lower for k in ['amount', 'sales', 'profit', 'cost', 'qty', 'rate', 'price']):
                    self.dimension_cols.append(col)
                else:
                    self.metric_cols.append(col)
            else:
                # String / Object column
                if unique_count <= min(total_rows * 0.85, 200) and unique_count > 0:
                    self.dimension_cols.append(col)
                else:
                    self.id_cols.append(col)
                    
        self.cleaned_df = df
        
        # Pick default primary metric (prefer names with Sales, Revenue, Amount, Profit, Total)
        metric_priority = ['sales', 'revenue', 'amount', 'profit', 'total', 'turnover', 'cost', 'qty', 'quantity', 'units', 'value', 'price']
        for prio in metric_priority:
            for m in self.metric_cols:
                if prio in m.lower():
                    self.primary_metric = m
                    break
            if self.primary_metric:
                break
        if not self.primary_metric and self.metric_cols:
            self.primary_metric = self.metric_cols[0]
            
        # Pick default primary dimension (prefer Region, Category, Product, Department, Status, Branch)
        dim_priority = ['category', 'region', 'department', 'product', 'branch', 'segment', 'status', 'channel', 'division', 'rep', 'manager']
        for prio in dim_priority:
            for d in self.dimension_cols:
                if prio in d.lower():
                    self.primary_dimension = d
                    break
            if self.primary_dimension:
                break
        if not self.primary_dimension and self.dimension_cols:
            self.primary_dimension = self.dimension_cols[0]
            
        # Pick primary date
        if self.date_cols:
            self.primary_date = self.date_cols[0]

    def get_summary_kpis(self, primary_metric: Optional[str] = None, df_subset: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
        """Compute high-level executive KPIs."""
        df = df_subset if df_subset is not None else self.cleaned_df
        if df is None or len(df) == 0:
            return {}

        metric = primary_metric or self.primary_metric
        total_records = len(df)
        
        kpis = {
            "total_records": total_records,
            "metric_name": metric,
            "total_value": 0.0,
            "avg_value": 0.0,
            "median_value": 0.0,
            "max_value": 0.0,
            "min_value": 0.0,
            "std_value": 0.0,
            "other_metrics": {},
            "date_range": None
        }
        
        if metric and metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]):
            series = df[metric].dropna()
            if len(series) > 0:
                kpis["total_value"] = float(series.sum())
                kpis["avg_value"] = float(series.mean())
                kpis["median_value"] = float(series.median())
                kpis["max_value"] = float(series.max())
                kpis["min_value"] = float(series.min())
                kpis["std_value"] = float(series.std())
                
        # Secondary metrics summary
        for m in self.metric_cols:
            if m != metric and m in df.columns and pd.api.types.is_numeric_dtype(df[m]):
                kpis["other_metrics"][m] = {
                    "total": float(df[m].sum()),
                    "avg": float(df[m].mean())
                }
                
        # Date range if date col exists
        if self.primary_date and self.primary_date in df.columns:
            dt_series = pd.to_datetime(df[self.primary_date], errors='coerce').dropna()
            if len(dt_series) > 0:
                kpis["date_range"] = {
                    "start": dt_series.min().strftime("%d-%b-%Y"),
                    "end": dt_series.max().strftime("%d-%b-%Y"),
                    "days": (dt_series.max() - dt_series.min()).days + 1
                }
                
        return kpis

    def get_dimension_breakdown(self, dimension_col: str, metric_col: Optional[str] = None, top_n: int = 10, df_subset: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """Calculate breakdown table for a given dimension with %, totals, and Pareto cumsum."""
        df = df_subset if df_subset is not None else self.cleaned_df
        metric = metric_col or self.primary_metric
        
        if df is None or dimension_col not in df.columns:
            return pd.DataFrame()
            
        if metric and metric in df.columns and pd.api.types.is_numeric_dtype(df[metric]):
            grp = df.groupby(dimension_col, dropna=False).agg(
                Total=(metric, 'sum'),
                Average=(metric, 'mean'),
                Count=(metric, 'count')
            ).reset_index()
            
            # Format null dimension names
            grp[dimension_col] = grp[dimension_col].fillna("Not Specified").astype(str)
            
            total_sum = grp['Total'].sum()
            grp['Share_Pct'] = (grp['Total'] / total_sum * 100).round(2) if total_sum > 0 else 0
            
            # Sort descending
            grp = grp.sort_values(by='Total', ascending=False).reset_index(drop=True)
            grp['Cumulative_Pct'] = (grp['Share_Pct'].cumsum()).round(2)
            
            # Group into Top N + Others if required
            if len(grp) > top_n:
                top_slice = grp.iloc[:top_n].copy()
                others_slice = grp.iloc[top_n:].copy()
                
                others_row = pd.DataFrame([{
                    dimension_col: "Others",
                    "Total": others_slice['Total'].sum(),
                    "Average": others_slice['Total'].sum() / others_slice['Count'].sum() if others_slice['Count'].sum() > 0 else 0,
                    "Count": others_slice['Count'].sum(),
                    "Share_Pct": round(others_slice['Share_Pct'].sum(), 2),
                    "Cumulative_Pct": 100.0
                }])
                return pd.concat([top_slice, others_row], ignore_index=True)
            return grp
        else:
            # Metric is not numeric or not available, do record frequency
            grp = df[dimension_col].fillna("Not Specified").value_counts().reset_index()
            grp.columns = [dimension_col, 'Count']
            grp['Share_Pct'] = (grp['Count'] / len(df) * 100).round(2)
            grp['Cumulative_Pct'] = (grp['Share_Pct'].cumsum()).round(2)
            if len(grp) > top_n:
                top_slice = grp.iloc[:top_n].copy()
                others_slice = grp.iloc[top_n:].copy()
                others_row = pd.DataFrame([{
                    dimension_col: "Others",
                    "Count": others_slice['Count'].sum(),
                    "Share_Pct": round(others_slice['Share_Pct'].sum(), 2),
                    "Cumulative_Pct": 100.0
                }])
                return pd.concat([top_slice, others_row], ignore_index=True)
            return grp

    def get_time_series_trend(self, date_col: Optional[str] = None, metric_col: Optional[str] = None, freq: str = "M", df_subset: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """Calculate time series aggregated trend with MoM (Month-over-Month) growth."""
        df = df_subset if df_subset is not None else self.cleaned_df
        d_col = date_col or self.primary_date
        m_col = metric_col or self.primary_metric
        
        if df is None or not d_col or d_col not in df.columns:
            return pd.DataFrame()
            
        temp_df = df[[d_col]].copy()
        temp_df['dt'] = pd.to_datetime(temp_df[d_col], errors='coerce')
        temp_df = temp_df.dropna(subset=['dt'])
        
        if len(temp_df) == 0:
            return pd.DataFrame()
            
        if m_col and m_col in df.columns and pd.api.types.is_numeric_dtype(df[m_col]):
            temp_df['val'] = df.loc[temp_df.index, m_col]
        else:
            temp_df['val'] = 1
            
        # Group by Period
        if freq == "M":
            temp_df['Period'] = temp_df['dt'].dt.to_period('M').astype(str)
        elif freq == "Q":
            temp_df['Period'] = temp_df['dt'].dt.to_period('Q').astype(str)
        else:
            temp_df['Period'] = temp_df['dt'].dt.to_period('Y').astype(str)
            
        trend = temp_df.groupby('Period').agg(
            Total=('val', 'sum'),
            Count=('val', 'count'),
            Avg=('val', 'mean')
        ).reset_index()
        
        trend = trend.sort_values(by='Period').reset_index(drop=True)
        # Calculate MoM growth %
        trend['Growth_Pct'] = (trend['Total'].pct_change() * 100).round(2).fillna(0.0)
        return trend

    def get_cross_tab(self, row_dim: str, col_dim: str, metric_col: Optional[str] = None, df_subset: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """Generate 2-way Pivot / Cross-Tabulation table."""
        df = df_subset if df_subset is not None else self.cleaned_df
        m_col = metric_col or self.primary_metric
        
        if df is None or row_dim not in df.columns or col_dim not in df.columns:
            return pd.DataFrame()
            
        if m_col and m_col in df.columns and pd.api.types.is_numeric_dtype(df[m_col]):
            pivot = pd.pivot_table(
                df,
                values=m_col,
                index=row_dim,
                columns=col_dim,
                aggfunc='sum',
                fill_value=0,
                margins=True,
                margins_name="Total"
            )
        else:
            pivot = pd.crosstab(
                df[row_dim],
                df[col_dim],
                margins=True,
                margins_name="Total"
            )
        return pivot

    def audit_data_quality(self) -> Dict[str, Any]:
        """Perform comprehensive data quality and exception checks."""
        df = self.cleaned_df
        if df is None:
            return {}
            
        total_rows = len(df)
        total_cells = df.size
        missing_cells = int(df.isna().sum().sum())
        completeness_pct = round((1 - (missing_cells / total_cells if total_cells > 0 else 0)) * 100, 2)
        
        # Missing values per column
        missing_by_col = []
        for col in df.columns:
            n_miss = int(df[col].isna().sum())
            if n_miss > 0:
                missing_by_col.append({
                    "Column": col,
                    "Missing_Count": n_miss,
                    "Missing_Pct": round(n_miss / total_rows * 100, 2)
                })
                
        # Duplicates check
        duplicates_count = int(df.duplicated().sum())
        
        # Outlier detection using IQR on numeric columns
        outliers_info = []
        for m in self.metric_cols:
            series = df[m].dropna()
            if len(series) > 10:
                q1 = series.quantile(0.25)
                q3 = series.quantile(0.75)
                iqr = q3 - q1
                lower_bound = q1 - 1.5 * iqr
                upper_bound = q3 + 1.5 * iqr
                outliers = series[(series < lower_bound) | (series > upper_bound)]
                if len(outliers) > 0:
                    outliers_info.append({
                        "Metric": m,
                        "Outlier_Count": int(len(outliers)),
                        "Outlier_Pct": round(len(outliers) / total_rows * 100, 2),
                        "Min_Outlier": float(outliers.min()),
                        "Max_Outlier": float(outliers.max()),
                        "Expected_Range": f"[{round(lower_bound, 2)}, {round(upper_bound, 2)}]"
                    })

        return {
            "total_rows": total_rows,
            "total_columns": len(df.columns),
            "missing_cells": missing_cells,
            "completeness_pct": completeness_pct,
            "missing_columns": missing_by_col,
            "duplicate_rows": duplicates_count,
            "outliers": outliers_info
        }

    def generate_executive_insights(self) -> List[str]:
        """Generate automated executive narrative bullet points."""
        insights = []
        kpis = self.get_summary_kpis()
        if not kpis:
            return ["No data available to generate insights."]

        metric_name = kpis.get("metric_name", "Records")
        tot = kpis.get("total_value", 0)
        cnt = kpis.get("total_records", 0)
        avg = kpis.get("avg_value", 0)

        # 1. High level scale
        if metric_name and tot > 0:
            insights.append(
                f"**Overall Scale**: Total {metric_name} is **{tot:,.2f}** across **{cnt:,}** records with an average of **{avg:,.2f}** per record."
            )
        else:
            insights.append(f"**Overall Scale**: Dataset contains **{cnt:,}** total transactions/records.")

        # 2. Date range span
        d_range = kpis.get("date_range")
        if d_range:
            daily_run_rate = (tot / d_range['days']) if d_range['days'] > 0 and tot > 0 else 0
            insights.append(
                f"**Time Horizon**: Data spans **{d_range['days']} days** (from {d_range['start']} to {d_range['end']})."
                + (f" Daily run-rate stands at **{daily_run_rate:,.2f}** per day." if daily_run_rate > 0 else "")
            )

        # 3. Top category contributor (Pareto observation)
        if self.primary_dimension:
            dim_df = self.get_dimension_breakdown(self.primary_dimension, top_n=5)
            if len(dim_df) > 0 and 'Share_Pct' in dim_df.columns:
                top_row = dim_df.iloc[0]
                top_name = top_row[self.primary_dimension]
                top_share = top_row['Share_Pct']
                top_val = top_row['Total'] if 'Total' in top_row else top_row['Count']
                insights.append(
                    f"**Dominant Driver ({self.primary_dimension})**: **'{top_name}'** is the leading segment, capturing **{top_share}%** share ({top_val:,.2f})."
                )

        # 4. Growth / Trend insight
        if self.primary_date:
            trend_df = self.get_time_series_trend(freq="M")
            if len(trend_df) >= 2:
                best_period = trend_df.sort_values(by='Total', ascending=False).iloc[0]
                latest_period = trend_df.iloc[-1]
                latest_growth = latest_period['Growth_Pct']
                growth_text = f"an increase of {latest_growth}%" if latest_growth >= 0 else f"a decline of {abs(latest_growth)}%"
                insights.append(
                    f"**Trend & Trajectory**: Highest activity was recorded in **{best_period['Period']}** ({best_period['Total']:,.2f}). Latest period ({latest_period['Period']}) showed {growth_text} MoM."
                )

        # 5. Data Health & Audit
        audit = self.audit_data_quality()
        if audit.get("duplicate_rows", 0) > 0:
            insights.append(f"**Attention**: Found **{audit['duplicate_rows']} duplicate records** that may require de-duplication.")
        if audit.get("missing_columns"):
            miss_count = len(audit["missing_columns"])
            insights.append(f"**Data Quality**: **{miss_count} columns** contain missing null values. Overall dataset completeness is **{audit.get('completeness_pct', 100)}%**.")

        return insights
