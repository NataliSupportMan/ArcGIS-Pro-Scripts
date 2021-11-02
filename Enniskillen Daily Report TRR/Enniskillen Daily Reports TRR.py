import pandas as pd
import xlsxwriter

# Adding the excel paths
df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonA_COL_Duct For TRR.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonB_COL_Duct For TRR.xlsx')
df3 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonC_COL_Duct For TRR.xlsx')
df4 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonD_COL_Duct For TRR.xlsx')
df5 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonE_COL_Duct For TRR.xlsx')
df6 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Enniskillen_Fibrus\STR_ENK1_RibbonF_COL_Duct For TRR.xlsx')


# Ribbon A Shape Length
group_DS1 = df1.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI1 = df1.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count1 = df1['Sub-Duct Installed?'] == 'Yes'
count1 = df1.loc[count1]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon A: ")
join1 = pd.concat([group_DS1, group_SDI1,  count1])
join1 = pd.DataFrame(data=join1)
join1 = join1.transpose()
join1 = join1[:1]
join1 = join1.drop(0, axis=1)
join1 = join1.drop('No', axis=1)
join1 = join1.rename(columns={'Yes': 'Sub-Ducted Sections'})
join1 = join1[['Congested Duct', 'Duct Blocked', 'Duct Clear', 'No Duct/Direct Burried', 'Other Notes (comments)', 'Sub-Ducted Sections']]
join1.insert(0, 'Status', 'WIP')
join1.insert(0, 'Enniskillen', 'Ribbon A')
join1.insert(7, 'Total Length Checked', '')
join1.insert(8, 'Cable Length raised for TRR (m)', '')
join1.insert(9, 'Remain', '')
join1.insert(10, '% Complete', '')
print(join1)

# Ribbon B Shape Length
group_DS2 = df2.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI2 = df2.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count2 = df2['Sub-Duct Installed?'] == 'Yes'
count2 = df2.loc[count2]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon B: ")
join2 = pd.concat([group_DS2, group_SDI2,  count2])
join2 = pd.DataFrame(data=join2)
join2 = join2.transpose()
join2 = join2.drop(0, axis=1)
join2 = join2.drop('No', axis=1)
join2 = join2.rename(columns={'Yes': 'Sub-Ducted Sections'})
join2.insert(0, 'Status', 'WIP')
join2.insert(0, 'Enniskillen', 'Ribbon B')
join2.insert(7, 'Total Length Checked', '')
join2.insert(8, 'Cable Length raised for TRR (m)', '')
join2.insert(9, 'Remain', '')
join2.insert(10, '% Complete', '')
print(join2)

# Ribbon C Shape Length
group_DS3 = df3.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI3 = df3.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count3 = df3['Sub-Duct Installed?'] == 'Yes'
count3 = df3.loc[count3]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon C: ")
join3 = pd.concat([group_DS3, group_SDI3, count3])
join3 = pd.DataFrame(data=join3)
join3 = join3.transpose()
join3 = join3.drop(0, axis=1)
join3 = join3.drop('No', axis=1)
join3 = join3.rename(columns={'Yes': 'Sub-Ducted Sections'})
join3.insert(0, 'Status', 'WIP')
join3.insert(0, 'Enniskillen', 'Ribbon C')
join3.insert(7, 'Total Length Checked', '')
join3.insert(8, 'Cable Length raised for TRR (m)', '')
join3.insert(9, 'Remain', '')
join3.insert(10, '% Complete', '')
print(join3)


# Ribbon D Shape Length
group_DS4 = df4.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI4 = df4.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count4 = df4['Sub-Duct Installed?'] == 'Yes'
count4 = df4.loc[count4]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon D: ")
join4 = pd.concat([group_DS4, group_SDI4, count4])
join4 = pd.DataFrame(data=join4)
join4 = join4.transpose()
join4 = join4.drop(0, axis=1)
join4 = join4.drop('No', axis=1)
join4 = join4.rename(columns={'Yes': 'Sub-Ducted Sections'})
join4.insert(0, 'Status', 'WIP')
join4.insert(0, 'Enniskillen', 'Ribbon D')
join4.insert(7, 'Total Length Checked', '')
join4.insert(8, 'Cable Length raised for TRR (m)', '')
join4.insert(9, 'Remain', '')
join4.insert(10, '% Complete', '')
print(join4)


# Ribbon E Shape Length
group_DS5 = df5.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI5 = df5.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count5 = df5['Sub-Duct Installed?'] == 'Yes'
count5 = df5.loc[count5]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon E: ")
join5 = pd.concat([group_DS5, group_SDI5, count5])
join5 = pd.DataFrame(data=join5)
join5 = join5.transpose()
join5 = join5.drop(0, axis=1)
join5 = join5.drop('No', axis=1)
join5 = join5.rename(columns={'Yes': 'Sub-Ducted Sections'})
join5.insert(0, 'Status', 'WIP')
join5.insert(0, 'Enniskillen', 'Ribbon E')
join5.insert(2, 'Congested Duct', '0')
join5.insert(7, 'Total Length Checked', '')
join5.insert(8, 'Cable Length raised for TRR (m)', '')
join5.insert(9, 'Remain', '')
join5.insert(10, '% Complete', '')
print(join5)


# Ribbon F Shape Length
group_DS6 = df6.groupby(['Duct Status '])['Shape__Length'].sum()
group_SDI6 = df6.groupby(['Sub-Duct Installed?'])['Shape__Length'].sum()
count6 = df6['Sub-Duct Installed?'] == 'Yes'
count6 = df6.loc[count6]['Sub-Duct Installed?'].value_counts()
print("Enniskillen Ribbon F: ")
join6 = pd.concat([group_DS6, group_SDI6, count6])
join6 = pd.DataFrame(data=join6)
join6 = join6.transpose()
join6 = join6.drop(0, axis=1)
join6 = join6.drop('No', axis=1)
join6 = join6.rename(columns={'Yes': 'Sub-Ducted Sections'})
join6.insert(0, 'Status', 'WIP')
join6.insert(0, 'Enniskillen', 'Ribbon F')
join6.insert(2, 'Congested Duct', '0')
join6.insert(7, 'Total Length Checked', '')
join6.insert(8, 'Cable Length raised for TRR (m)', '')
join6.insert(9, 'Remain', '')
join6.insert(10, '% Complete', '')
print(join6)

# Concat all the tables/rows into 1 table
df_all = pd.concat([join1, join2, join3, join4, join5, join6], ignore_index=True)
print(df_all.to_string())


# Covert the DataFrame to xlsx writer
writer = pd.ExcelWriter('Enniskillen Daily Reports TRR.xlsx', engine='xlsxwriter')
df_all.to_excel(writer, sheet_name='Enniskillen')
workbook = writer.book
worksheet = writer.sheets['Enniskillen']


# Add a header format
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'fg_color': '#171717',
    'font_color': 'white',
    'valign': 'vcenter',
})

# Write the column headers with the defined format.
for col_num, value in enumerate(df_all.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)

# Adding the formats
format1 = workbook.add_format({'bg_color': '#FFE699', 'font_color': '#000000'})
format2 = workbook.add_format({'bg_color': '#D7E4BC', 'font_color': '#000000'})
bor_format = workbook.add_format({'border': 2})
per_format = workbook.add_format({'num_format': '0.00%'})
comma_format = workbook.add_format({'num_format': '####,#'})

# Calculate the columns to find the sum
worksheet.write('I2', '{=SUM(D2:H2)}')
worksheet.write('I3', '{=SUM(D3:H3)}')
worksheet.write('I4', '{=SUM(D4:H4)}')
worksheet.write('I5', '{=SUM(D5:H5)}')
worksheet.write('I6', '{=SUM(D6:H6)}')
worksheet.write('I7', '{=SUM(D7:H7)}')

# Adding the total length checked from TRR
worksheet.write('J2', 17810)
worksheet.write('J3', 37558)
worksheet.write('J4', 34509)
worksheet.write('J5', 21598)
worksheet.write('J6', 20562)
worksheet.write('J7', 23083)

# Calculating the remain and adding the format
worksheet.write('K2', '=J2 - I2', comma_format)
worksheet.write('K3', '=J3 - I3', comma_format)
worksheet.write('K4', '=J4 - I4', comma_format)
worksheet.write('K5', '=J5 - I5', comma_format)
worksheet.write('K6', '=J6 - I6', comma_format)
worksheet.write('K7', '=J7 - I7', comma_format)

# Calculating the percentage and adding format
worksheet.write('L2', '=I2 / J2 ', per_format)
worksheet.write('L3', '=I3 / J3 ', per_format)
worksheet.write('L4', '=I4 / J4 ', per_format)
worksheet.write('L5', '=I5 / J5 ', per_format)
worksheet.write('L6', '=I6 / J6 ', per_format)
worksheet.write('L7', '=I7 / J7 ', per_format)

# Adding the color on the specific rows
worksheet.conditional_format('B2:B2', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B3:B3', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B4:B4', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B5:B5', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B6:B6', {'type': 'unique', 'format': format2})
worksheet.conditional_format('B7:B7', {'type': 'unique', 'format': format2})

# Adding the color on the specific rows
worksheet.conditional_format('C2:C2', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C3:C3', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C4:C4', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C5:C5', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C6:C6', {'type': 'unique', 'format': format1})
worksheet.conditional_format('C7:C7', {'type': 'unique', 'format': format1})

# Adding the border format for the cells
worksheet.conditional_format('B2:N2', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B3:N3', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B4:N4', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B5:N5', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B6:N6', {'type': 'unique', 'format': bor_format})
worksheet.conditional_format('B7:N7', {'type': 'unique', 'format': bor_format})

# Saving the worksheet to local folder
writer.save()