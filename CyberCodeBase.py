# This code was created by Kiana Lashgari and Daniel Jusino. We can be contacted at Kiana_Lashgari@ajgre.com and Daniel_Jusino@ajgre.com

# This tool depends on the output of the Cyber Data Prep Tool. Specifically the Prism-Re Input Columns Template

# Step 1 - Import the dataset that includes Public Sector and/or Private Households into the Python environment. Change the file path below!

from operator import index
import pandas as pd

# CHANGE FILE PATH:
data = pd.read_excel("C:\Kiana\Python testing.xlsx")

# Exposure summary:
total_prem = (data['Policy Premium at Participation'].sum())
gross_prem = round(total_prem/1e6,2)
ID_count = (data['Insured ID'].count())
avg_lim = round((data['Breach 1st party\nCoverage Limit\nat 100%\n($)'].mean())/1e6,2)
avg_att = round((data['Breach 1st party\nCoverage\nAttachment / Deductible\n($)'].mean())/1e6,2)
avg_rev = round((data['Annual Revenue\n($M)'].mean())/1e6,2)

exp_table = {'GWP ($M)': [gross_prem], 'Count': [ID_count], 'Average Limit($M)': [avg_lim], 'Average Attachment($M)': [avg_att], 'Average Revenue($M)':[avg_rev]}
exposure_summary = pd.DataFrame(exp_table, columns= ['GWP ($M)', 'Count', 'Average Limit($M)', 'Average Attachment($M)', 'Average Revenue($M)'])

# Sector summary:
sectors = set(data['PRISM-Re Industry (breach)'])
sector_summary = pd.DataFrame(data.groupby('PRISM-Re Industry (breach)')['Policy Premium at Participation'].sum()/total_prem)

# CHANGE FILE PATH:
with pd.ExcelWriter('C:\Kiana\Python output.xlsx', engine='xlsxwriter') as writer:
    exposure_summary.to_excel(writer, sheet_name='Summary', index=False)
    sector_summary.to_excel(writer, sheet_name='Summary', index=True, startrow=4)