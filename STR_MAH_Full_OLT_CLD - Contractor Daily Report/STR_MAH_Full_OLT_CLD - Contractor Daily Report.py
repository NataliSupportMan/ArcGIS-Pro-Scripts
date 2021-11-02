import pandas as pd

df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Poles.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Chambers.xlsx')
df3 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Fibre Cable.xlsx')
df4 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Fibre Duct.xlsx')
df5 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Splice Closures.xlsx')
df6 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\Maghera Vegetation Issues.xlsx')

# Convert the tree_lengt from string to numeric
df6 = df6.fillna(0)
df6['tree_lengt'] = df6.tree_lengt.astype(str).str.extract('(^\d*)').astype(float)

# Poles table added and filtered by Built
poles_filter = df1['Status'] == 'Built'
poles = df1.loc[poles_filter]['Status'].value_counts()
poles = poles.rename(index={'Built': 'Poles'})

# Chambers tabled added and filtered by Built
chambers_filter = df2['Status'] == 'Built'
chambers = df2.loc[chambers_filter]['Status'].value_counts()
chambers = chambers.rename(index={'Built': 'Chambers'})

# Fibre cable added and grouped by Placement
fibre_cables = df3.loc[df3['Status'] == 'Built', ['Placement', 'Shape__Length']].groupby('Placement').sum()

# Fibre duct table added and grouped by Status
fibre_duct = df4.loc[df4['Status'] == 'Built', ['Status', 'Shape__Length']].groupby('Status').sum()
fibre_duct = fibre_duct.rename(index={'Built': 'Civils'})

# Splice closures added and grouped by Enclosure Type
splice_closures = df5.loc[df5['Status'] == 'Spliced', ['Enclosure Type', 'Status']].groupby('Enclosure Type').count()

# Vegetation table added and groupd by Status
vegetation = df6.loc[df6['Status'] == 'Completed', ['Status', 'tree_lengt']].groupby('Status').sum()
vegetation = vegetation.rename(index={'Completed': 'Tree Cutting'})


# Concatinating all the tables to one table and adding all the columns to one column
df_all = pd.concat([vegetation, splice_closures, poles, chambers, fibre_duct, fibre_cables])
df_all = df_all.fillna(0) # Filling all the NAN with 0 to avoid errors
df_all.loc[:, 'QTY COMPLETE'] = df_all[0] + df_all['Shape__Length'] + df_all['tree_lengt'] + df_all['Status'] # Adding all the columns to one column
df_all = df_all.reset_index() # Reset the index of the table
df_all = df_all.rename(columns={'index': 'TASK'}) # Rename the index to TASK
df_all = df_all.drop(columns=[0, 'Shape__Length', 'tree_lengt', 'Status']) # Dropping the no needed columns
df_all['BOM QTY'] = ([32112.8, 108, 38, 141, 576, 19, 886, 32, 7202.25, 345041, 181518.67, 256997.75])
df_all = df_all[['TASK', 'BOM QTY', 'QTY COMPLETE']]
print(df_all)

# Transfer the Database to Excel writer to export the dataframe
writer = pd.ExcelWriter('STR_MAH_Full_OLT_CLD - Contractor Daily Report.xlsx', engine='xlsxwriter',)

# Adding a variable
sheet_name = 'Maghera Build'

# Setting the rows and the columns
df_all.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False)

# Adding a workbook to alter the tables
workbook = writer.book
worksheet = writer.sheets[sheet_name]
bold = workbook.add_format({'bold': True})

# Format the cells
cell_format = workbook.add_format({'bold': True, 'font_size': '14',
                                   'font_name': 'Calibri Light',
                                   'valign': 'vcenter', 'text_wrap': True})
# Format the Header
header_format = workbook.add_format({'bold': True, 'font_size': '17',
                                    'font_name': 'Calibri Light',
                                    'valign': 'vcenter'})

#Adding the header
worksheet.write_string(0, 1, 'Maghera Build', header_format)

# Adding a bigger size for header
worksheet.set_row(0, 30,)

# Adding a bigger size for the columns headers
worksheet.set_row(1, 18,)

# Adding a color format
format1 = workbook.add_format({'bg_color': '#DCE6F1', 'font_color': '#000000'}) # Light blue
format2 = workbook.add_format({'bg_color': '#D5E2B8', 'font_color': '#000000'}) # Light olive
format3 = workbook.add_format({'bg_color': '#ffffff', 'font_color': '#000000'}) # Light grey

# Declare columns with color light blue
worksheet.conditional_format('B1:B1', {'type': 'unique', 'format': format3})
worksheet.conditional_format('B2:B14', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C4:C4', {'type': 'unique', 'format': format2})
worksheet.conditional_format('C6:C6', {'type': 'unique', 'format': format2})
worksheet.conditional_format('C8:C8', {'type': 'unique', 'format': format2})
worksheet.conditional_format('C10:C10', {'type': 'unique', 'format': format2})
worksheet.conditional_format('C2:C14', {'type': 'unique', 'format': format2})

# Set a list for the widths
colwidths = {}

# Store the defaults.
for col in range(50000):
    colwidths[col] = 15

# Calculate the width manually
colwidths[0] = 12
colwidths[1] = 10
colwidths[2] = 16

# Then set the column widths.
for col_num, width in colwidths.items():
    worksheet.set_column(col_num, col_num, width)

# Saving to the local folder
writer.save()

