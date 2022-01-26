# This code was created by Kiana Lashgari and Daniel Jusino. We can be contacted at Kiana_Lashgari@ajgre.com and Daniel_Jusino@ajgre.com

# This tool depends on the output of the Cyber Data Prep Tool. Specifically the Prism-Re Input Columns Template

# Step 1 - Import the dataset that includes Public Sector and/or Private Households into the Python environment. Change the file path below!

from operator import index
import pandas as pd

# Change file path:
data = pd.read_excel("C:\Kiana\Python testing.xlsx")

gross_prem = round((data['Policy Premium at Participation'].sum())/1e6,2)
ID_count = (data['Insured ID'].count())
avg_lim = round((data['Breach 1st party\nCoverage Limit\nat 100%\n($)'].mean())/1e6,2)
avg_att = round((data['Breach 1st party\nCoverage\nAttachment / Deductible\n($)'].mean())/1e6,2)
avg_rev = round((data['Annual Revenue\n($M)'].mean())/1e6,2)

results_summary = {'GWP ($M)': [gross_prem], 'Count': [ID_count], 'Average Limit($M)': [avg_lim], 'Average Attachment($M)': [avg_att], 'Average Revenue($M)':[avg_rev]}

df = pd.DataFrame(results_summary, columns= ['GWP ($M)', 'Count', 'Average Limit($M)', 'Average Attachment($M)', 'Average Revenue($M)'])

# Change file path and choose output excel name:
df.to_excel('C:\Kiana\Python output.xlsx', index= False, header = True)