import pandas as pd
import xlsxwriter
import openpyxl

#This helps me see the whole df instead of truncated version
pd.set_option('display.max_rows', None, 'display.max_columns', None)

#This opens the excel sheet we want to update and turns into a df. I also had to remove the "_" from some of the headers for the code to work.
df = pd.read_excel("biosamples.147.2022-11-14T10-14-18.xlsx")

#here I opened al the .tsv turn them into df files and assigned a variable to use later
date_df = pd.read_csv('samn_collection_date_update.tsv', sep='\t')
lat_ion_df = pd.read_csv('samn_isolate_lat_lon_update.tsv', sep='\t')
isolation_source_df = pd.read_csv('samn_isolation_source_update.tsv', sep='\t')

#samn_isolate_update.tsv is not used because the information is the same as samn_isolate_lat_ion_update.tsv (used excel to compare duplicate values)

#This updates the main df using the .update function. (Look for pandas API reference on .update funtions)
df.update(date_df)
df.update(lat_ion_df)
df.update(isolation_source_df)

#This saves the sheet and retains the other sheets in the workbook
with pd.ExcelWriter("biosamples.147.2022-11-14T10-14-18.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, 'Biosamples', index=False)