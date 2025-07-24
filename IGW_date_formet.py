# ==== Main() ====
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# Hide the root Tkinter window
Tk().withdraw()

# Ask for required files
file_path = askopenfilename(title="Select IGW_TRAFFIC_TREND File", filetypes=[("Excel Files", "*.xlsx")])
file_path_N1 = askopenfilename(title="Select Output File Name", filetypes=[("Excel Files", "*.xlsx")])

# Load Excel file
df_xls = pd.ExcelFile(file_path)
df = pd.read_excel(file_path, sheet_name=df_xls.sheet_names[0])

# Use actual date column
date_column = 'D_DATE'

# Convert to datetime format
df[date_column] = pd.to_datetime(df[date_column])

# Format date1 and Minute
df['date1'] = df[date_column].dt.strftime('%d-%b-%Y')
df['Minute'] = df[date_column].dt.minute

# IGW_Date formatting (Windows and Linux support)
try:
    df['IGW_Date'] = df[date_column].dt.strftime('%a %b %#d %H:%M:%S EAT %Y')  # Windows
except:
    df['IGW_Date'] = df[date_column].dt.strftime('%a %b %-d %H:%M:%S EAT %Y')  # Linux/Mac

# Save to Excel
with pd.ExcelWriter(file_path_N1) as writer:
    df.to_excel(writer, sheet_name='data', index=False)

print("✅ Successfully completed and saved!")
