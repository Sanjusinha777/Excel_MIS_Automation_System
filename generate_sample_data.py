"""
Sample Data Generator for Excel MIS Automation
Creates realistic enterprise business sample data (Sales, Operations, Finance)
"""

import os
import random
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

def generate_sales_mis_data(num_rows=500, filename="Sample_Sales_MIS.xlsx"):
    random.seed(42)
    np.random.seed(42)

    regions = ["North", "South", "East", "West", "Central"]
    branches = {
        "North": ["Delhi", "Chandigarh", "Jaipur", "Lucknow"],
        "South": ["Bengaluru", "Chennai", "Hyderabad", "Kochi"],
        "East": ["Kolkata", "Bhubaneswar", "Patna", "Guwahati"],
        "West": ["Mumbai", "Pune", "Ahmedabad", "Surat"],
        "Central": ["Bhopal", "Indore", "Nagpur", "Raipur"]
    }
    categories = {
        "Electronics": ["Laptop", "Smartphone", "Tablet", "Headphones", "Smartwatch"],
        "Furniture": ["Office Chair", "Executive Desk", "Bookshelf", "Filing Cabinet"],
        "Office Supplies": ["Printer Paper", "Ink Cartridge", "Stapler", "Notepad", "Pens Pack"]
    }
    channels = ["Corporate Direct", "Online Portal", "Retail Partner", "Government Tender"]
    statuses = ["Completed", "Completed", "Completed", "Pending", "Cancelled"]

    start_date = datetime(2025, 1, 1)
    
    rows = []
    for i in range(1, num_rows + 1):
        order_id = f"ORD-2025-{1000 + i}"
        days_offset = random.randint(0, 365)
        order_date = start_date + timedelta(days=days_offset)
        
        region = random.choice(regions)
        branch = random.choice(branches[region])
        
        cat = random.choice(list(categories.keys()))
        prod = random.choice(categories[cat])
        
        channel = random.choice(channels)
        status = random.choice(statuses)
        
        # Quantities and base prices
        if cat == "Electronics":
            qty = random.randint(1, 15)
            unit_price = round(random.uniform(5000, 75000), 2)
            margin_rate = random.uniform(0.12, 0.28)
        elif cat == "Furniture":
            qty = random.randint(1, 20)
            unit_price = round(random.uniform(3000, 35000), 2)
            margin_rate = random.uniform(0.18, 0.35)
        else:
            qty = random.randint(5, 100)
            unit_price = round(random.uniform(200, 2500), 2)
            margin_rate = random.uniform(0.25, 0.45)
            
        gross_sales = round(qty * unit_price, 2)
        discount_pct = random.choice([0.0, 0.05, 0.10, 0.15, 0.20])
        discount_amount = round(gross_sales * discount_pct, 2)
        net_sales = round(gross_sales - discount_amount, 2)
        
        cost_of_goods = round(net_sales * (1 - margin_rate), 2)
        gross_profit = round(net_sales - cost_of_goods, 2)
        
        # Sales Representative
        reps = ["Rahul Sharma", "Priya Nair", "Amit Verma", "Sneha Patel", "Vikram Singh", "Ananya Roy"]
        rep = random.choice(reps)
        
        rows.append({
            "Order ID": order_id,
            "Date": order_date.strftime("%Y-%m-%d"),
            "Sales Rep": rep,
            "Region": region,
            "Branch": branch,
            "Category": cat,
            "Product": prod,
            "Sales Channel": channel,
            "Quantity": qty,
            "Unit Price (INR)": unit_price,
            "Gross Sales (INR)": gross_sales,
            "Discount Amount (INR)": discount_amount,
            "Net Sales (INR)": net_sales,
            "Cost (INR)": cost_of_goods,
            "Gross Profit (INR)": gross_profit,
            "Order Status": status
        })
        
    df = pd.DataFrame(rows)
    
    # Introduce a couple of realistic anomalies (e.g., slight missing values for testing data cleaning)
    df.loc[15, "Branch"] = None
    df.loc[42, "Discount Amount (INR)"] = None
    
    output_path = os.path.join(os.path.dirname(__file__), filename) if "__file__" in locals() else filename
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sales_Transactions", index=False)
        
        # Second sheet for targets / quotas
        targets = []
        for reg in regions:
            for rep in ["Rahul Sharma", "Priya Nair", "Amit Verma", "Sneha Patel", "Vikram Singh", "Ananya Roy"]:
                targets.append({
                    "Region": reg,
                    "Sales Rep": rep,
                    "Annual Target (INR)": random.randint(2500000, 6000000)
                })
        pd.DataFrame(targets).to_excel(writer, sheet_name="Targets", index=False)
        
    print(f"Sample MIS Excel dataset created successfully at: {output_path}")
    return output_path

if __name__ == "__main__":
    generate_sales_mis_data(500, "Sample_Sales_MIS.xlsx")
