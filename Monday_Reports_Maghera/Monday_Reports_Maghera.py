import pandas as pd
import numpy as np


df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonA_COL_Duct For TRR.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonB_COL_Duct For TRR.xlsx')
df3 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonC_COL_Duct For TRR.xlsx')
df4 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonD_COL_Duct For TRR.xlsx')
df5 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonE_COL_Duct For TRR.xlsx')
df6 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonF_COL_Duct For TRR.xlsx')
df7 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Maghera_Fibrus\STR_MAH1_RibbonG_COL_Duct For TRR.xlsx')

#Ribbon A
#Add a new columns and re-order the headers
df1 = df1[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df1.insert(5, 'No. of Sections', 1)
df1 = df1.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values1 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df1 = df1.fillna(value=values1)

# Creating First pivot table
pivot_ribbon_a1 = pd.pivot_table(new_df1, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_a1 = pivot_ribbon_a1.stack('Duct Status ')
pivot_ribbon_a1 = pivot_ribbon_a1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_a2 = pd.pivot_table(new_df1, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_a2 = pivot_ribbon_a2.stack('Sub-Duct Installed?')
pivot_ribbon_a2 = pivot_ribbon_a2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon B
#Add a new columns and re-order the headers
df2 = df2[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df2.insert(5, 'No. of Sections', 1)
df2 = df2.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values2 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df2 = df2.fillna(value=values2)

# Creating First pivot table
pivot_ribbon_b1 = pd.pivot_table(new_df2, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_b1 = pivot_ribbon_b1.stack('Duct Status ')
pivot_ribbon_b1 = pivot_ribbon_b1[['Sum Shape Length', 'Sum No. of Sections']].copy()

# Creating Second pivot table
pivot_ribbon_b2 = pd.pivot_table(new_df2, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_b2 = pivot_ribbon_b2.stack('Sub-Duct Installed?')
pivot_ribbon_b2 = pivot_ribbon_b2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon C
#Add a new columns and re-order the headers
df3 = df3[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df3.insert(5, 'No. of Sections', 1)
df3 = df3.rename(columns={'Shape__Length': 'Sum Shape Length',
                         'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values3 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df3 = df3.fillna(value=values3)

# Creating First pivot table
pivot_ribbon_c1 = pd.pivot_table(new_df3, values=['Sum Shape Length', 'Sum No. of Sections'],
                                 index=['TRR Crew'], columns=['Duct Status '],
                                 aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_c1 = pivot_ribbon_c1.stack('Duct Status ')
pivot_ribbon_c1 = pivot_ribbon_c1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_c2 = pd.pivot_table(new_df3, values=['Sum Shape Length', 'Sum No. of Sections'],
                                 index=['TRR Crew'],
                                 columns=['Sub-Duct Installed?'],
                                 aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_c2 = pivot_ribbon_c2.stack('Sub-Duct Installed?')
pivot_ribbon_c2 = pivot_ribbon_c2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon D
#Add a new columns and re-order the headers
df4 = df4[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df4.insert(5, 'No. of Sections', 1)
df4 = df4.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values4 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df4 = df4.fillna(value=values4)

# Creating First pivot table
pivot_ribbon_d1 = pd.pivot_table(new_df4, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_d1 = pivot_ribbon_d1.stack('Duct Status ')
pivot_ribbon_d1 = pivot_ribbon_d1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_d2 = pd.pivot_table(new_df4, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_d2 = pivot_ribbon_d2.stack('Sub-Duct Installed?')
pivot_ribbon_d2 = pivot_ribbon_d2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon E
#Add a new columns and re-order the headers
df5 = df5[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df5.insert(5, 'No. of Sections', 1)
df5 = df5.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values5 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df5 = df5.fillna(value=values5)

# Creating First pivot table
pivot_ribbon_e1 = pd.pivot_table(new_df5, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_e1 = pivot_ribbon_e1.stack('Duct Status ')
pivot_ribbon_e1 = pivot_ribbon_e1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_e2 = pd.pivot_table(new_df5, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_e2 = pivot_ribbon_e2.stack('Sub-Duct Installed?')
pivot_ribbon_e2 = pivot_ribbon_e2[['Sum Shape Length', 'Sum No. of Sections']].copy()

#Ribbon F
#Add a new columns and re-order the headers
df6 = df6[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df6.insert(5, 'No. of Sections', 1)
df6 = df6.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values6 = {'TRR Crew': 'No TRR Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df6 = df6.fillna(value=values6)

# Creating First pivot table
pivot_ribbon_f1 = pd.pivot_table(new_df6, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_f1 = pivot_ribbon_f1.stack('Duct Status ')
pivot_ribbon_f1 = pivot_ribbon_f1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_f2 = pd.pivot_table(new_df6, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_f2 = pivot_ribbon_f2.stack('Sub-Duct Installed?')
pivot_ribbon_f2 = pivot_ribbon_f2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon G
#Add a new columns and re-order the headers
df7 = df7[['TRR Crew', 'Sub Ducting Crew', 'Blockage Clearance Crew', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df7.insert(5, 'No. of Sections', 1)
df7 = df7.rename(columns={'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values7 = {'TRR Crew': 'No Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df7 = df7.fillna(value=values6)

# Creating First pivot table
pivot_ribbon_g1 = pd.pivot_table(new_df7, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_g1 = pivot_ribbon_g1.stack('Duct Status ')
pivot_ribbon_g1 = pivot_ribbon_g1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_g2 = pd.pivot_table(new_df7, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_g2 = pivot_ribbon_g2.stack('Sub-Duct Installed?')
pivot_ribbon_g2 = pivot_ribbon_g2[['Sum Shape Length', 'Sum No. of Sections']].copy()


# We add all the tables together
# Concat all the ribbons to one table
df_all = pd.concat([new_df1, new_df2, new_df3, new_df4, new_df5, new_df6, new_df7])

# Creating First pivot table
all_pivots_1 = pd.pivot_table(df_all, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
all_pivots_1 = all_pivots_1.stack('Duct Status ')
all_pivots_1 = all_pivots_1[['Sum Shape Length', 'Sum No. of Sections']].copy()

# Creating Second pivot table
all_pivots_2 = pd.pivot_table(df_all, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['TRR Crew'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
all_pivots_2 = all_pivots_2.stack('Sub-Duct Installed?')
all_pivots_2 = all_pivots_2[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Convert all the DataFrames to xlsxwriter
writer = pd.ExcelWriter('Maghera.xlsx', engine='xlsxwriter')

# The DataFrames with pivot tables converted to writer
df_all.to_excel(writer, sheet_name='All Ribbons', index=False)
all_pivots_1.to_excel(writer, sheet_name='All Ribbons', startrow=3, startcol=9)
all_pivots_2.to_excel(writer, sheet_name='All Ribbons', startrow=3, startcol=14)

new_df1.to_excel(writer, sheet_name='Ribbon A Duct For TRR', index=False)
pivot_ribbon_a1.to_excel(writer, sheet_name='Ribbon A Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_a2.to_excel(writer, sheet_name='Ribbon A Duct For TRR', startrow=3, startcol=14)

new_df2.to_excel(writer, sheet_name='Ribbon B Duct For TRR', index=False)
pivot_ribbon_b1.to_excel(writer, sheet_name='Ribbon B Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_b2.to_excel(writer, sheet_name='Ribbon B Duct For TRR', startrow=3, startcol=14)

new_df3.to_excel(writer, sheet_name="Ribbon C Duct For TRR", index=False)
pivot_ribbon_c1.to_excel(writer, sheet_name='Ribbon C Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_c2.to_excel(writer, sheet_name='Ribbon C Duct For TRR', startrow=3, startcol=14)

new_df4.to_excel(writer, sheet_name='Ribbon D Duct For TRR', index=False)
pivot_ribbon_d1.to_excel(writer, sheet_name='Ribbon D Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_d2.to_excel(writer, sheet_name='Ribbon D Duct For TRR', startrow=3, startcol=14)

new_df5.to_excel(writer, sheet_name='Ribbon E Duct For TRR', index=False)
pivot_ribbon_e1.to_excel(writer, sheet_name='Ribbon E Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_e2.to_excel(writer, sheet_name='Ribbon E Duct For TRR', startrow=3, startcol=14)

new_df6.to_excel(writer, sheet_name='Ribbon F Duct For TRR', index=False)
pivot_ribbon_f1.to_excel(writer, sheet_name='Ribbon F Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_f2.to_excel(writer, sheet_name='Ribbon F Duct For TRR', startrow=3, startcol=14)

new_df7.to_excel(writer, sheet_name='Ribbon G Duct For TRR', index=False)
pivot_ribbon_g1.to_excel(writer, sheet_name='Ribbon G Duct For TRR', startrow=3, startcol=9)
pivot_ribbon_g2.to_excel(writer, sheet_name='Ribbon G Duct For TRR', startrow=3, startcol=14)

# Creating a workbook adding the sheets
workbook = writer.book
worksheet_all = writer.sheets['All Ribbons']
worksheet1 = writer.sheets['Ribbon A Duct For TRR']
worksheet2 = writer.sheets['Ribbon B Duct For TRR']
worksheet3 = writer.sheets['Ribbon C Duct For TRR']
worksheet4 = writer.sheets['Ribbon D Duct For TRR']
worksheet5 = writer.sheets['Ribbon E Duct For TRR']
worksheet6 = writer.sheets['Ribbon F Duct For TRR']
worksheet7 = writer.sheets['Ribbon G Duct For TRR']

# Adding colour format
format1 = workbook.add_format({'bg_color': '#DCE6F1', 'font_color': '#000000'})

# All ribbons headers colour
worksheet_all.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet_all.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet_all.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon A headers colour
worksheet1.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet1.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet1.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon B headers colour
worksheet2.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet2.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet2.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon C headers colour
worksheet3.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet3.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet3.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon D headers colour
worksheet4.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet4.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet4.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon E headers colour
worksheet5.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet5.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet5.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon F headers colour
worksheet6.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet6.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet6.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Ribbon G headers colour
worksheet7.conditional_format('A1:H1', {'type': 'unique', 'format': format1})
worksheet7.conditional_format('J4:M4', {'type': 'unique', 'format': format1})
worksheet7.conditional_format('O4:R4', {'type': 'unique', 'format': format1})

# Adding auto filter to all worksheets headers
worksheet_all.autofilter('A1:H1')
worksheet_all.autofilter(0, 0, 5000, 7)

worksheet1.autofilter('A1:H1')
worksheet1.autofilter(0, 0, 5000, 7)

worksheet2.autofilter('A1:H1')
worksheet2.autofilter(0, 0, 5000, 7)

worksheet3.autofilter('A1:H1')
worksheet3.autofilter(0, 0, 5000, 7)

worksheet4.autofilter('A1:H1')
worksheet4.autofilter(0, 0, 5000, 7)

worksheet5.autofilter('A1:H1')
worksheet5.autofilter(0, 0, 5000, 7)

worksheet6.autofilter('A1:H1')
worksheet6.autofilter(0, 0, 5000, 7)

worksheet7.autofilter('A1:H1')
worksheet7.autofilter(0, 0, 5000, 7)

# Set a list for the widths
colwidths = {}

# Store the defaults.
for col in range(50000):
    colwidths[col] = 15

# Calculate the width manually
colwidths[0] = 20
colwidths[1] = 22
colwidths[2] = 24
colwidths[3] = 7
colwidths[4] = 20
colwidths[5] = 23
colwidths[6] = 15
colwidths[7] = 22
colwidths[8] = 5
colwidths[9] = 18
colwidths[10] = 24
colwidths[11] = 19
colwidths[12] = 19
colwidths[13] = 5
colwidths[14] = 18
colwidths[15] = 25
colwidths[16] = 18
colwidths[17] = 18

# Then set the column widths.
for col_num, width in colwidths.items():
    worksheet_all.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet1.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet2.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet3.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet4.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet5.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet6.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet7.set_column(col_num, col_num, width)

writer.save()
