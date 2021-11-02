import pandas as pd


df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Garrison_Fibrus\STR_GRS1_RibbonA_COL_Duct For TRR.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Garrison_Fibrus\STR_GRS1_RibbonB_COL_Duct For TRR.xlsx')


cols = ['Congested Duct',  'Duct Blocked', 'Duct Clear',
        'No Duct/Direct Burried', 'Other Notes (comments)', 'Sub-Duct Installed?', 'Sub-Duct Meterage']

# Ribbon A
new_df1 = df1[['Duct Status ', 'Shape__Length_3', 'Sub-Duct Installed?']]
df1.loc[df1['Duct Status '] == 0, 'Duct Status '] = ''
ribbon_a = new_df1.groupby(['Duct Status '])['Shape__Length_3'].sum()
ribbon_a_sub = new_df1.groupby(['Sub-Duct Installed?'])['Shape__Length_3'].sum()
count_sub_a = new_df1['Sub-Duct Installed?'] == 'Yes'
count_a = new_df1.loc[count_sub_a]['Sub-Duct Installed?'].value_counts()
concat_a = pd.concat([ribbon_a, ribbon_a_sub, count_a])
new_ribbon_a = concat_a.reset_index()
new_ribbon_a = new_ribbon_a.rename(columns={'index': ' ', 0: 'Ribbon A'})
new_ribbon_a = new_ribbon_a.transpose()
headers_a = new_ribbon_a.iloc[0]
new_ribbon_a = pd.DataFrame(new_ribbon_a.values[1:], columns=headers_a)
new_ribbon_a.columns = [x[1] if x[1] not in new_ribbon_a.columns[:x[0]]
                        else f"{x[1]}_{list(new_ribbon_a.columns[:x[0]]).count(x[1])}"
                        for x in enumerate(new_ribbon_a.columns)]
new_ribbon_a = new_ribbon_a.rename(columns={'Yes': 'Sub-Duct Meterage', 'Yes_1': 'Sub-Duct Installed?'})
new_ribbon_a = new_ribbon_a.rename(index={0: 0})
del new_ribbon_a[0]
#del new_ribbon_a['0_1']
del new_ribbon_a['No']
new_ribbon_a = new_ribbon_a.reindex(new_ribbon_a.columns.union(cols), axis=1, fill_value=0)
new_ribbon_a.insert(0, 'Status', 'WIP')
new_ribbon_a.insert(0, 'Garrison', 'Ribbon A')
new_ribbon_a.insert(7, 'Total Length Checked', '')
new_ribbon_a.insert(8, 'Cable Length raised for TRR (m)', '')
new_ribbon_a.insert(9, 'Remain', '')
new_ribbon_a.insert(10, '% Complete', '')
new_ribbon_a.insert(13, 'Assisted Digs', '')
new_ribbon_a.insert(14, 'Desilts Sections', '')
new_ribbon_a.insert(15, 'Desilts Meterage', '')
print(new_ribbon_a.to_string())


# Ribbon B
new_df2 = df2[['Duct Status ', 'Shape__Length_3', 'Sub-Duct Installed?']]
df2.loc[df2['Duct Status '] == 0, 'Duct Status '] = ''
ribbon_b = new_df2.groupby(['Duct Status '])['Shape__Length_3'].sum()
ribbon_b_sub = new_df2.groupby(['Sub-Duct Installed?'])['Shape__Length_3'].sum()
count_b_sub = new_df2['Sub-Duct Installed?'] == 'Yes'
count_b = new_df2.loc[count_b_sub]['Sub-Duct Installed?'].value_counts()
concat_b = pd.concat([ribbon_b, ribbon_b_sub, count_b])
new_ribbon_b = concat_b.reset_index()
new_ribbon_b = new_ribbon_b.rename(columns={'index': ' ', 0: 'Ribbon B'})
new_ribbon_b = new_ribbon_b.transpose()
headers_b = new_ribbon_b.iloc[0]
new_ribbon_b = pd.DataFrame(new_ribbon_b.values[1:], columns=headers_b)
new_ribbon_b.columns = [x[1] if x[1] not in new_ribbon_b.columns[:x[0]]
                        else f"{x[1]}_{list(new_ribbon_b.columns[:x[0]]).count(x[1])}"
                        for x in enumerate(new_ribbon_b.columns)]
new_ribbon_b = new_ribbon_b.rename(columns={'Yes': 'Sub-Duct Meterage', 'Yes_1': 'Sub-Duct Installed?'})
new_ribbon_b = new_ribbon_b.rename(index={0: 1})
del new_ribbon_b[0]
del new_ribbon_b['0_1']
del new_ribbon_b['No']
new_ribbon_b = new_ribbon_b.reindex(new_ribbon_b.columns.union(cols), axis=1, fill_value=0)
new_ribbon_b.insert(0, 'Status', 'WIP')
new_ribbon_b.insert(0, 'Garrison', 'Ribbon B')
new_ribbon_b.insert(7, 'Total Length Checked', '')
new_ribbon_b.insert(8, 'Cable Length raised for TRR (m)', '')
new_ribbon_b.insert(9, 'Remain', '')
new_ribbon_b.insert(10, '% Complete', '')
new_ribbon_b.insert(13, 'Assisted Digs', '')
new_ribbon_b.insert(14, 'Desilts Sections', '')
new_ribbon_b.insert(15, 'Desilts Meterage', '')
print(new_ribbon_b.to_string())

# Concat all the tables into one table
new_table = pd.concat([new_ribbon_a, new_ribbon_b])
print(new_table.to_string())


# Covert the DataFrame to xlsx writer
writer = pd.ExcelWriter('Garrison Daily Report TRR.xlsx', engine='xlsxwriter')
new_table.to_excel(writer, sheet_name='Garrison')
workbook = writer.book
worksheet = writer.sheets['Garrison']

# Add a header format
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'fg_color': '#171717',
    'font_color': 'white',
    'valign': 'vcenter',
})

# Write the column headers with the defined format.
for col_num, value in enumerate(new_table.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)

# Adding the various formats, color, borders, percentages
format1 = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#000000'})
format2 = workbook.add_format({'bg_color': '#D7E4BC', 'font_color': '#000000'})
bor_format = workbook.add_format({'border': 2})
per_format = workbook.add_format({'num_format': '0.00%'})
comma_format = workbook.add_format({'num_format': '####,#'})

# Sum up the columns
worksheet.write('I2', '{=SUM(D2:H2)}')
worksheet.write('I3', '{=SUM(D3:H3)}')

# Adding the total length checked for TRR
worksheet.write('J2', 14568)
worksheet.write('J3', 26496)

# Calculating the remain and adding format
worksheet.write('K2', '=J2 - I2', comma_format)
worksheet.write('K3', '=J3 - I3', comma_format)

# calculating the percentage and adding the format
worksheet.write('L2', '=I2 / J2 ', per_format)
worksheet.write('L3', '=I3 / J3 ', per_format)

# Adding color format to specific rows
worksheet.conditional_format('B2:B2', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B3:B3', {'type': 'unique', 'format': format2})

# Adding color format to specific rows
worksheet.conditional_format('C2:C2', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C3:C3', {'type': 'unique', 'format': format1})

# Adding borders format to specific rows
worksheet.conditional_format('B2:N2', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B3:N3', {'type': 'unique', 'format': bor_format})

# Saving the worksheet to local file
writer.save()