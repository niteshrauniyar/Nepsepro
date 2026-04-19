"""
generate_nepse_sample.py
Generates a realistic NEPSE-style floorsheet for testing.
Run: python generate_nepse_sample.py
Output: nepse_sample_floorsheet.xlsx
"""

import pandas as pd
import numpy as np

np.random.seed(99)

STOCKS = {
    "NABIL": {"base": 1250, "vol_mult": 1.4},
    "SCB":   {"base": 823,  "vol_mult": 1.0},
    "NICA":  {"base": 945,  "vol_mult": 1.2},
    "NMB":   {"base": 475,  "vol_mult": 0.9},
    "ADBL":  {"base": 612,  "vol_mult": 1.1},
    "GBIME": {"base": 389,  "vol_mult": 0.8},
    "HIDCL": {"base": 274,  "vol_mult": 0.7},
    "NRIC":  {"base": 1845, "vol_mult": 0.6},
}

# Broker pools
INSTITUTIONAL = [12, 34, 56, 23, 41]
RETAIL        = list(range(100, 145))
ALL_BROKERS   = INSTITUTIONAL + RETAIL

records = []
sn = 1

for symbol, cfg in STOCKS.items():
    n_trades = np.random.randint(80, 220)
    base     = cfg["base"]
    vm       = cfg["vol_mult"]

    # Decide stock character
    bias = np.random.choice(["accumulation","distribution","neutral"])

    price = base
    for i in range(n_trades):
        drift = np.random.normal(0, base * 0.003)
        price = round(max(price + drift, base * 0.85), 2)

        rand = np.random.random()

        if bias == "accumulation" and rand < 0.20:
            buyer  = np.random.choice(INSTITUTIONAL[:2])
            seller = np.random.choice(RETAIL)
            qty    = int(np.random.normal(900*vm, 80))
        elif bias == "distribution" and rand < 0.18:
            buyer  = np.random.choice(RETAIL)
            seller = np.random.choice(INSTITUTIONAL[:2])
            qty    = int(np.random.normal(1100*vm, 100))
        elif rand < 0.04:
            buyer  = np.random.choice(ALL_BROKERS)
            seller = np.random.choice(ALL_BROKERS)
            qty    = int(np.random.normal(5000*vm, 400))  # block trade
        else:
            buyer  = np.random.choice(ALL_BROKERS)
            seller = np.random.choice(ALL_BROKERS)
            while seller == buyer:
                seller = np.random.choice(ALL_BROKERS)
            qty = int(np.random.exponential(400 * vm))

        qty = max(10, qty)
        amount = round(qty * price, 2)

        records.append({
            "SN":               sn,
            "Contract No":      f"C{sn:06d}",
            "Stock Symbol":     symbol,
            "Buyer Broker No":  buyer,
            "Seller Broker No": seller,
            "Quantity":         qty,
            "Rate":             price,
            "Amount":           amount,
        })
        sn += 1

df = pd.DataFrame(records)
df.to_excel("nepse_sample_floorsheet.xlsx", index=False)

print(f"✅ Generated nepse_sample_floorsheet.xlsx")
print(f"   Total trades : {len(df):,}")
print(f"   Symbols      : {df['Stock Symbol'].nunique()} — {list(df['Stock Symbol'].unique())}")
print(f"   Unique buyers: {df['Buyer Broker No'].nunique()}")
print(f"   Total amount : Rs {df['Amount'].sum():,.2f}")
print(f"\nSample rows:")
print(df.head(5).to_string(index=False))
