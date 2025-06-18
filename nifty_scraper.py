import os
import pandas as pd
from datetime import datetime
from nsepython import option_chain
from openpyxl.chart import LineChart, Reference

# 1) Fetch & prepare data once
raw = option_chain("NIFTY")["records"]["data"]
df = pd.DataFrame(raw)
df["expiryDate"] = pd.to_datetime(df["expiryDate"], dayfirst=True)

# 2) Isolate front-month
front = df["expiryDate"].min()
df_front = df[df["expiryDate"] == front]

# 3) Build IV DataFrame
df_iv = pd.DataFrame({
    "StrikePrice": df_front["strikePrice"],
    "ExpiryDate":  df_front["expiryDate"],
    "PE_IV":       df_front["PE"].map(lambda x: x.get("impliedVolatility") if isinstance(x, dict) else None),
    "CE_IV":       df_front["CE"].map(lambda x: x.get("impliedVolatility") if isinstance(x, dict) else None)
}).sort_values("StrikePrice")

# 4) Write everything in one go and embed a native Excel chart
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename  = f"nifty_dashboard_{timestamp}.xlsx"
with pd.ExcelWriter(filename, engine="openpyxl") as writer:
    # Write sheets
    df_front.to_excel(writer, sheet_name="Raw Data", index=False)
    df_iv.to_excel(writer, sheet_name="IV Data",  index=False)

    # Access workbook and IV Data sheet
    book = writer.book
    ws   = book["IV Data"]

    # Create Excel native LineChart for volatility skew
    chart = LineChart()
    chart.title = f"Volatility Skew for {front.date()}"
    chart.x_axis.title = "Strike Price"
    chart.y_axis.title = "Implied Vol (%)"

    max_row = ws.max_row
    # StrikePrice in column A (1), PE_IV in C (3), CE_IV in D (4)
    cats    = Reference(ws, min_col=1, min_row=2, max_row=max_row)
    pe_data = Reference(ws, min_col=3, min_row=1, max_row=max_row)
    ce_data = Reference(ws, min_col=4, min_row=1, max_row=max_row)

    chart.add_data(pe_data, titles_from_data=True)
    chart.add_data(ce_data, titles_from_data=True)
    chart.set_categories(cats)

    # Position chart
    ws.add_chart(chart, "F2")

print(f"Created {filename} with Raw Data, IV Data, and embedded native skew chart.")
