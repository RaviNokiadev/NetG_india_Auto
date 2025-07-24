import pandas as pd
from datetime import timedelta
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Step 1: Load File
Tk().withdraw()
file_path = askopenfilename(title="Select Output File", filetypes=[("Excel Files", "*.xlsx")])
df = pd.read_excel(file_path, sheet_name='data')

# Step 2: Preprocess
df['D_DATE'] = pd.to_datetime(df['D_DATE']).dt.floor('min')
df = df[df['IGW_NAME'].isin(['Kenya IGW', 'Djibouti IGW'])]

# Step 3: Time slots
start_time = df['D_DATE'].min().replace(second=0, microsecond=0)
end_time = df['D_DATE'].max()
time_slots = pd.date_range(start=start_time, end=end_time, freq='15T')

# Step 4: Prepare Output
rows = []

for t in time_slots:
    row = {'Time': t.strftime('%a %b %d %H:%M:%S EAT %Y')}
    comments = []

    for igw in ['Kenya IGW', 'Djibouti IGW']:
        mask = (df['IGW_NAME'] == igw) & (df['D_DATE'] >= t) & (df['D_DATE'] <= t + timedelta(minutes=2))
        match = df.loc[mask]

        if not match.empty:
            actual_time = match.iloc[0]['D_DATE']
            max_value = match['MAX'].max()  # ← FIXED HERE
            row[igw] = round(max_value, 4)
            if actual_time != t:
                comments.append(f"{igw.split()[0]}→{actual_time.strftime('%H:%M')}")
        else:
            row[igw] = 'Missing'
            comments.append(f"{igw.split()[0]}→Missing")
    
    row['Comments'] = ', '.join(comments)
    rows.append(row)

# Step 5: Export
output_df = pd.DataFrame(rows)
output_file = file_path.replace('.xlsx', '_15min_MAX_Kenya_Djibouti.xlsx')

with pd.ExcelWriter(output_file) as writer:
    output_df.to_excel(writer, sheet_name='15min_MAX', index=False)

print("✅ File Created:", output_file)
