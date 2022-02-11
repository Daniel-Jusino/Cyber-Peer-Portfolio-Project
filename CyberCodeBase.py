# This code was created by Kiana Lashgari and Daniel Jusino. We can be contacted at Kiana_Lashgari@ajgre.com and Daniel_Jusino@ajgre.com

# This tool depends on the output of the Cyber Data Prep Tool. Specifically the Prism-Re Input Columns Template

# MAKE SURE TO Import the dataset that includes Public Sector with revised Private Households sectors into the Python environment. Change the file path below!

from operator import index
import pandas as pd
import numpy as np

# CHANGE FILE PATH:
data = pd.read_excel("C:\Kiana\Python testing.xlsx")

# Exposure summary:
total_prem = (data['Policy Premium at Participation'].sum())
gross_prem = round(total_prem/1e6,2)
ID_count = (data['Insured ID'].count())
avg_lim = round((data['Breach 1st party\nCoverage Limit\nat 100%\n($)'].mean())/1e6,2)
avg_att = round((data['Breach 1st party\nCoverage\nAttachment / Deductible\n($)'].mean())/1e6,2)
avg_rev = round((data['Annual Revenue\n($M)'].mean())/1e6,2)
avg_security = round((data['Cybersecurity Level (1 = Above Average, 2 = Average, 3 = Below Average)'].mean()),2)

exp_table = {'GWP ($M)': [gross_prem], 'Count': [ID_count], 'Average Limit ($M)': [avg_lim], 'Average Attachment ($M)': [avg_att], 'Average Revenue ($M)':[avg_rev], 'Average Cyber Security Level':[avg_security]}
exposure_summary = pd.DataFrame(exp_table, columns= ['GWP ($M)', 'Count', 'Average Limit ($M)', 'Average Attachment ($M)', 'Average Revenue ($M)', 'Average Cyber Security Level'])

# Sector summary:
sectors = set(data['PRISM-Re Industry (breach)'])
sector_summary = pd.DataFrame(data.groupby('PRISM-Re Industry (breach)')['Policy Premium at Participation'].sum()/total_prem)

# Revenue summary:
bins = [-1, 9.999999, 49.999999, 99.999999, 249.999999, 499.999999, 999.999999, 1999.999999, np.inf]
labels = ['Less than $10M', '$10M - $50M', '$50M - $100M', '$100M - $250M', '$250M - $500M', '$500M - $1B', '$1B - $2B', 'Greater than $2B']
rev_band = data.groupby(pd.cut(data['Annual Revenue\n($M)'], bins = bins , labels = labels))
rev_band_sum = round(rev_band['Annual Revenue\n($M)'].sum(),2)

# CHANGE FILE PATH:
with pd.ExcelWriter('C:\Kiana\Python output.xlsx', engine='xlsxwriter') as writer:
    exposure_summary.to_excel(writer, sheet_name='Summary', index=False)
    sector_summary.to_excel(writer, sheet_name='Summary', index=True, startrow=4)
    rev_band_sum.to_excel(writer, sheet_name='Summary', index=True, startrow=21)