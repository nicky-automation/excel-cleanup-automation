import pandas as pd
from datetime import datetime

# Load the messy Excel file
df = pd.read_excel("messy_customer_data.xlsx")

# --- CLEANING STEPS ---

# 1. Trim spaces from all string columns
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# 2. Standardize Customer Name capitalization
df['Customer Name'] = df['Customer Name'].str.title()

# 3. Clean and format Amount column
df['Amount'] = (
    df['Amount']
    .astype(str)
    .str.replace('£', '', regex=False)
    .str.replace(',', '', regex=False)
    .str.replace(' ', '', regex=False)
    .astype(float)
)
df['Amount (£)'] = df['Amount'].apply(lambda x: f"£{x:,.2f}")
df.drop(columns=['Amount'], inplace=True)

# 4. Standardize Status capitalization
df['Status'] = df['Status'].str.strip().str.capitalize()

# 5. Convert mixed date formats to datetime objects
def parse_date(x):
    for fmt in ("%d/%m/%y", "%Y-%m-%d", "%B %d %Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(x), fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return x

df['Date'] = df['Date'].apply(parse_date)

# 6. Sort by Date
df = df.sort_values(by='Date', ascending=True)

# 7. Save to new Excel file
df.to_excel("cleaned_customer_data.xlsx", index=False)

print("✅ Cleaning complete! 'cleaned_customer_data.xlsx' has been created.")
