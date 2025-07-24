from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import sys

# Hide the Tkinter root window
Tk().withdraw()

# Ask for file paths
file_path_R1 = askopenfilename(title="Select CORE_LINK_TRAFFIC_TREND (.xlsb)", filetypes=[("Excel Binary Workbook", "*.xlsb")])
file_path_R2 = askopenfilename(title="Select AGG_LINK_TRAFFIC_TREND (.xlsb)", filetypes=[("Excel Binary Workbook", "*.xlsb")])
file_path_M1 = askopenfilename(title="Select Capacity Master file (.xlsx)", filetypes=[("Excel Files", "*.xlsx")])
file_path_N1 = askopenfilename(title="Select combine output File (.xlsx)", filetypes=[("Excel Files", "*.xlsx")])
" \
"

# put File (.xlsx)", filetypes=[("Excel Files", "*.xlsx")])

# Helper function to get the sheet containing "data"
def get_sheet_name_containing(file_path, keyword):
    with pd.ExcelFile(file_path, engine='pyxlsb') as xls:
        for sheet in xls.sheet_names:
            if keyword.lower() in sheet.lower():
                return sheet
        raise ValueError(f"No sheet with '{keyword}' found in {file_path}")

# Load both .xlsb files
sheet_R1 = get_sheet_name_containing(file_path_R1, "data")
sheet_R2 = get_sheet_name_containing(file_path_R2, "data")
dfR1 = pd.read_excel(file_path_R1, sheet_name=sheet_R1, engine='pyxlsb')
dfR2 = pd.read_excel(file_path_R2, sheet_name=sheet_R2, engine='pyxlsb')

# Combine
df = pd.concat([dfR1, dfR2], ignore_index=True)
print(f"✅ Combined Rows after merge: {len(df)}")

# Normalize columns
df.columns = df.columns.str.strip().str.upper()

# Check required columns
required_columns = {'SITE_NAME', 'PORT', 'D_DATE', 'MAX_TRAFFIC'}
if not required_columns.issubset(df.columns):
    print("❌ Error: Missing required columns in raw data.")
    print("Found columns:", df.columns.tolist())
    sys.exit()

# Handle date/time conversion safely
df['D_DATE'] = pd.to_datetime(df['D_DATE'], errors='coerce')

# If still bad, handle Excel serial numbers manually
if df['D_DATE'].isna().all() or df['D_DATE'].min().year == 1970:
    df_raw = pd.concat([dfR1, dfR2], ignore_index=True)
    df_raw.columns = df_raw.columns.str.strip().str.upper()
    df['D_DATE'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(df_raw['D_DATE'].astype(float), unit='D')

# Extract date/time parts
df['DATE1'] = df['D_DATE'].dt.normalize()
df['WEEKDAY'] = df['D_DATE'].dt.day_name().str[:3]
df['MONTH'] = df['D_DATE'].dt.month_name().str[:3]
df['MINUTE'] = df['D_DATE'].dt.minute
df['DAY'] = df['D_DATE'].dt.day
df['TIME_ONLY'] = df['D_DATE'].dt.time

# Create merge key
df['SITENAME_PORT'] = (
    df['SITE_NAME'].astype(str).str.strip().str.upper() +
    df['PORT'].astype(str).str.strip().str.upper()
)

# Load Capacity Master
dfM1 = pd.read_excel(file_path_M1)
dfM1.columns = dfM1.columns.str.strip().str.upper()
dfM1['SITENAME_PORT'] = dfM1['SITENAME_PORT'].astype(str).str.strip().str.upper()

# Check expected columns
expected_capacity_columns = {'SITENAME_PORT', 'CAPACITY_NEW'}
if not expected_capacity_columns.issubset(dfM1.columns):
    print("❌ Error: Missing expected columns in Capacity Master.")
    print("Found:", dfM1.columns.tolist())
    sys.exit()

# Merge with capacity
df1 = pd.merge(df, dfM1[['SITENAME_PORT', 'CAPACITY_NEW']], on='SITENAME_PORT', how='left')

# Find unmatched
unmatched_df = df1[df1['CAPACITY_NEW'].isna()][['SITE_NAME', 'PORT', 'SITENAME_PORT']].drop_duplicates()
if not unmatched_df.empty:
    print("⚠️ Warning: Some ports couldn't match with Capacity Master.")
    print("Example unmatched entries:")
    print(unmatched_df.head())

# Utilisation
df1['UTILISATION'] = (df1['MAX_TRAFFIC'] * 100) / df1['CAPACITY_NEW']
df1['UTILISATION'] = df1['UTILISATION'].round(2)

# Pivot
df1 = df1.sort_values(by='D_DATE')
dfP = pd.pivot_table(df1, values='UTILISATION',
                     index='SITENAME_PORT',
                     columns='DATE1',
                     aggfunc='max')

# Max 7-day utilisation
dfP['MAX_7D'] = dfP.max(axis=1)

# Ranges
bins = [0, 60, 70, 80, 90, 1000]
labels = ['Less than 60', 'Between 60-70', 'Between 70-80', 'Between 80-90', 'More than 90']
dfP['RANGE'] = pd.cut(dfP['MAX_7D'], bins=bins, labels=labels, right=False)

# Save to Excel
with pd.ExcelWriter(file_path_N1) as writer:
    df1.to_excel(writer, sheet_name='data', index=False)
    dfP.to_excel(writer, sheet_name='pivot', index=True)
    unmatched_df.to_excel(writer, sheet_name='unmatched', index=False)

print("✅ Completed: All sheets saved to Excel.")
