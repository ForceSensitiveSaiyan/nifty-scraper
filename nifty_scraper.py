import os
import pandas as pd
from datetime import datetime
from io import BytesIO
from nsepython import option_chain
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl import load_workbook

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

# 4) Create skew plot in-memory
buf = BytesIO()
plt.figure(figsize=(8,4))
plt.plot(df_iv["StrikePrice"], df_iv["CE_IV"], label="Call IV")
plt.plot(df_iv["StrikePrice"], df_iv["PE_IV"], label="Put IV")
plt.xlabel("Strike Price"); plt.ylabel("Implied Vol (%)")
plt.title(f"Vol Skew for {front.date()}")
plt.legend(); plt.grid(); plt.tight_layout()
plt.savefig(buf, format="png")
plt.close()
buf.seek(0)

# 5) Write everything in one go
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename  = f"nifty_dashboard_{timestamp}.xlsx"
with pd.ExcelWriter(filename, engine="openpyxl") as writer:
    # Write sheets
    df_front.to_excel(writer, sheet_name="Raw Data", index=False)
    df_iv.to_excel(writer, sheet_name="IV Data",  index=False)

    # Embed the image before closing
    book = writer.book
    ws   = book["IV Data"]
    img  = Image(buf)
    img.anchor = "F2"
    ws.add_image(img)

print(f"Created {filename} with Raw Data, IV Data, and embedded skew chart.")
