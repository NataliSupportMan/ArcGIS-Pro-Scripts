import pandas as pd
import numpy as np
import xlsxwriter


df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Warrenpoint_Fibrus\STR_WRP1_RibbonA_COL_Duct_For_TRR.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Warrenpoint_Fibrus\STR_WRP1_RibbonB_COL_Duct_For_TRR.xlsx')


#Ribbon A
#Add a new columns and re-order the headers
df1 = df1[['Editor', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df1.insert(5, 'No. of Sections', 1)
df1 = df1.rename(columns={'Editor': 'Crew breakdown', 'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values1 = {'Crew breakdown': 'No Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df1 = df1.fillna(value=values1)

# Creating First pivot table
pivot_ribbon_a1 = pd.pivot_table(new_df1, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_a1 = pivot_ribbon_a1.stack('Duct Status ')
pivot_ribbon_a1 = pivot_ribbon_a1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
pivot_ribbon_a2 = pd.pivot_table(new_df1, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_a2 = pivot_ribbon_a2.stack('Sub-Duct Installed?')
pivot_ribbon_a2 = pivot_ribbon_a2[['Sum Shape Length', 'Sum No. of Sections']].copy()


#Ribbon B
#Add a new columns and re-order the headers
df2 = df2[['Editor', 'id', 'Shape__Length', 'Duct Status ', 'Sub-Duct Installed?']]
df2.insert(5, 'No. of Sections', 1)
df2 = df2.rename(columns={'Editor': 'Crew breakdown', 'Shape__Length': 'Sum Shape Length',
                          'No. of Sections': 'Sum No. of Sections'})

# Filling the NaN values with Blank
values2 = {'Crew breakdown': 'No Crew', 'Duct Status ': 'No Duct Status',
           'Sub-Duct Installed?': 'No Sub-Duct Installed?', 'Sum Shape Length': 0}
new_df2 = df2.fillna(value=values2)

# Creating First pivot table
pivot_ribbon_b1 = pd.pivot_table(new_df2, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_b1 = pivot_ribbon_b1.stack('Duct Status ')
pivot_ribbon_b1 = pivot_ribbon_b1[['Sum Shape Length', 'Sum No. of Sections']].copy()

# Creating Second pivot table
pivot_ribbon_b2 = pd.pivot_table(new_df2, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
pivot_ribbon_b2 = pivot_ribbon_b2.stack('Sub-Duct Installed?')
pivot_ribbon_b2 = pivot_ribbon_b2[['Sum Shape Length', 'Sum No. of Sections']].copy()

# We add all the tables together
# Concat all the ribbons to one table
df_all = pd.concat([new_df1, new_df2])

# Creating First pivot table
all_pivots_1 = pd.pivot_table(df_all, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Duct Status '],
                        aggfunc=np.sum, margins=True, margins_name='All')
all_pivots_1 = all_pivots_1.stack('Duct Status ')
all_pivots_1 = all_pivots_1[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Creating Second pivot table
all_pivots_2 = pd.pivot_table(df_all, values=['Sum Shape Length', 'Sum No. of Sections'],
                        index=['Crew breakdown'], columns=['Sub-Duct Installed?'],
                        aggfunc=np.sum, margins=True, margins_name='All')
all_pivots_2 = all_pivots_2.stack('Sub-Duct Installed?')
all_pivots_2 = all_pivots_2[['Sum Shape Length', 'Sum No. of Sections']].copy()


# Convert all the DataFrames to xlsxwriter
writer = pd.ExcelWriter('Warrenpoint.xlsx', engine='xlsxwriter')

# The DataFrames with pivot tables converted to writer
df_all.to_excel(writer, sheet_name='All Ribbons', index=False)
all_pivots_1.to_excel(writer, sheet_name='All Ribbons', startrow=3, startcol=7,)
all_pivots_2.to_excel(writer, sheet_name='All Ribbons', startrow=3, startcol=12,)

new_df1.to_excel(writer, sheet_name='Ribbon A Duct For TRR', index=False)
pivot_ribbon_a1.to_excel(writer, sheet_name='Ribbon A Duct For TRR', startrow=3, startcol=7,)
pivot_ribbon_a2.to_excel(writer, sheet_name='Ribbon A Duct For TRR', startrow=3, startcol=12,)

new_df2.to_excel(writer, sheet_name='Ribbon B Duct For TRR', index=False)
pivot_ribbon_b1.to_excel(writer, sheet_name='Ribbon B Duct For TRR', startrow=3, startcol=7,)
pivot_ribbon_b2.to_excel(writer, sheet_name='Ribbon B Duct For TRR', startrow=3, startcol=12,)

# Creating a workbook adding the sheets
workbook = writer.book
worksheet_all = writer.sheets['All Ribbons']
worksheet1 = writer.sheets['Ribbon A Duct For TRR']
worksheet2 = writer.sheets['Ribbon B Duct For TRR']

# Adding colour format
format1 = workbook.add_format({'bg_color': '#DCE6F1', 'font_color': '#000000'})

# All ribbons headers colour
worksheet_all.conditional_format('A1:F1', {'type': 'unique', 'format': format1})
worksheet_all.conditional_format('H4:K4', {'type': 'unique', 'format': format1})
worksheet_all.conditional_format('M4:P4', {'type': 'unique', 'format': format1})

# Ribbon A headers colour
worksheet1.conditional_format('A1:F1', {'type': 'unique', 'format': format1})
worksheet1.conditional_format('H4:K4', {'type': 'unique', 'format': format1})
worksheet1.conditional_format('M4:P4', {'type': 'unique', 'format': format1})

# Ribbon B headers colour
worksheet2.conditional_format('A1:F1', {'type': 'unique', 'format': format1})
worksheet2.conditional_format('H4:K4', {'type': 'unique', 'format': format1})
worksheet2.conditional_format('M4:P4', {'type': 'unique', 'format': format1})

# Adding auto filter to all worksheets headers
worksheet_all.autofilter('A1:F1')
worksheet_all.autofilter(0, 0, 5000, 5)

worksheet1.autofilter('A1:F1')
worksheet1.autofilter(0, 0, 5000, 5)

worksheet2.autofilter('A1:F1')
worksheet2.autofilter(0, 0, 5000, 5)

# Set a list for the widths
colwidths = {}

# Store the defaults.
for col in range(50000):
    colwidths[col] = 15

# Calculate the width manually
colwidths[0] = 20
colwidths[1] = 7
colwidths[2] = 22
colwidths[3] = 23
colwidths[4] = 22
colwidths[5] = 23
colwidths[6] = 6
colwidths[7] = 16
colwidths[8] = 22
colwidths[9] = 18
colwidths[10] = 19
colwidths[11] = 5
colwidths[12] = 16
colwidths[13] = 22
colwidths[14] = 18
colwidths[15] = 18

# Then set the column widths.
for col_num, width in colwidths.items():
    worksheet_all.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet1.set_column(col_num, col_num, width)

for col_num, width in colwidths.items():
    worksheet2.set_column(col_num, col_num, width)

writer.save()