import pandas as pd
import numpy as np
import xlsxwriter

df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\Maghera Vegetation Issues.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Poles.xlsx')
df3 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Fibre Duct.xlsx')
df4 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Chambers.xlsx')
df5 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Fibre Cable.xlsx')
df6 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\STR MAH1 Full OLT CLD - Contractor - Splice Closures.xlsx')


# Convert the tree_lengt from string to numeric
df1['tree_lengt'] = pd.to_numeric(df1['tree_lengt'], errors='coerce')

# Filling the NaN values to 0
df1 = df1.fillna(0)

# Vegetation DataFrame to prepare the table
df1 = df1[['Milestone', 'tree_lengt', 'Status']] # Veg
group_mil1 = df1.groupby(['Milestone'])['tree_lengt'].agg(['count'])
tc_com = df1.loc[df1['Status'] == 'Completed', ['tree_lengt', 'Milestone']].groupby('Milestone').sum()
total_tree = pd.merge(group_mil1, tc_com, right_index=True, left_index=True, how="outer")
total_tree = total_tree.rename(columns={'count': 'TC BOM', 'tree_lengt': 'TC COM'})

# Poles DataFrame to prepare the table
df2 = df2[['Milestone', 'Status']] # Poles
group_mil2 = df2.loc[df2['Status'] == 'Planned', ['Status', 'Milestone']].groupby('Milestone').count()
sum_pole = df2.loc[df2['Status'] == 'Built', ['Milestone', 'Status']].groupby('Milestone').count()
total_poles = pd.merge(group_mil2, sum_pole, right_index=True, left_index=True, how="outer")
total_poles = total_poles.rename(columns={'Status_x': 'Pole BOM', 'Status_y': 'Pole COM'})

# Fibre Duct to prepare the table
df3 = df3[['Milestone', 'Status', 'Shape__Length']] # Fibre duct
group_mil3 = df3.loc[df3['Status'] == 'Planned', ['Shape__Length', 'Milestone']].groupby('Milestone').sum()
sum_fibre_duct = df3['Status'] == 'Built'
sum_fibre_count = df3.loc[sum_fibre_duct]['Milestone'].value_counts()
total_fibre_duct = pd.merge(group_mil3, sum_fibre_count, right_index=True, left_index=True, how="outer")
total_fibre_duct = total_fibre_duct.rename(columns={'Shape__Length': 'Civils BOM', 'Milestone': 'Civils COM'})

# Chambers to prepare the table
df4 = df4[['Milestone', 'Status']] # Chambers
group_mil4 = df4.loc[df4['Status'] == 'Planned', ['Status', 'Milestone']].groupby('Milestone').count()
chambers = df4.loc[df4['Status'] == 'Built', ['Status', 'Milestone']].groupby('Milestone').count()
total_chambers = pd.merge(group_mil4, chambers, right_index=True, left_index=True, how="outer")
total_chambers = total_chambers.rename(columns={'Status_x': 'Chambers BOM', 'Status_y': 'Chambers COM'})

# Fibre Cable to prepare the table of OH
df5 = df5[['Milestone', 'Status', 'Placement', 'Shape__Length']] # Fibre cable
group_oh = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'OH')].groupby('Milestone').sum()
total_oh = df5.loc[(df5['Placement'] == 'OH') & (df5['Status'] == 'Built'), ['Shape__Length', 'Milestone']].groupby('Milestone').sum()
total_fibre_oh = pd.merge(group_oh, total_oh, right_index=True, left_index=True, how="left")
total_fibre_oh = total_fibre_oh.rename(columns={'Shape__Length_x': 'O/H BOM', 'Shape__Length_y': 'O/H COM'})

# Fibre Cable to prepare the table of UG
group_ug = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'UG')].groupby('Milestone').sum()
total_ug = df5.loc[(df5['Placement'] == 'UG') & (df5['Status'] == 'Built'), ['Shape__Length', 'Milestone']].groupby('Milestone').sum()
total_fibre_ug = pd.merge(group_ug, total_ug, right_index=True, left_index=True, how="left")
total_fibre_ug = total_fibre_ug.rename(columns={'Shape__Length_x': 'U/G BOM', 'Shape__Length_y': 'U/G COM'})

# Fibre Cable to prepare the table of UG/OH
group_ug_oh = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'UG/OH')].groupby('Milestone').sum()
total_oh_ug = df5.loc[(df5['Placement'] == 'UG/OH') & (df5['Status'] == 'Built'), ['Shape__Length', 'Milestone']].groupby('Milestone').sum()
total_fibre_oh_ug = pd.merge(group_ug_oh, total_oh_ug, right_index=True, left_index=True, how="outer")
total_fibre_oh_ug = total_fibre_oh_ug.rename(columns={'Shape__Length_x': 'UG/OH BOM', 'Shape__Length_y': 'UG/OH COM'})

# ODP Splice Closures to prepare the table with condition
df6 = df6[['Milestone', 'Status', 'Enclosure Type']]
ODP = ['ODP1', 'ODP2', 'ODP3', 'ODP1/ODP2'] # Adding rows of calculation
filt_odp = (df6['Status'] == 'Planned') & (df6['Enclosure Type'].isin(ODP)) # Creating a filter to add the ODP
group_odp1 = df6.loc[filt_odp].groupby('Milestone').count() # Adding the filter and groupby Milestone
group_odp1 = group_odp1.rename(columns={'Enclosure Type': 'ODP BOM'}) # Rename columns
group_odp1 = group_odp1.drop(columns=['Status']) # Drop columns

# ODP Splice Closure preparing the table with condition
ODP2 = ['ODP1', 'ODP2', 'ODP3', 'ODP1/ODP2'] # Adding rows of calculation
filt_odp2 = (df6['Status'] == 'Spliced') & (df6['Enclosure Type'].isin(ODP2)) # Creating a filter to add the ODP
group_odp2 = df6.loc[filt_odp2].groupby('Milestone').count() # Adding the filter and groupby Milestone
group_odp2 = group_odp2.rename(columns={'Enclosure Type': 'ODP COM'}) # Rename columns
group_odp2 = group_odp2.drop(columns=['Status']) # Drop columns
total_odp = pd.merge(group_odp1, group_odp2, right_index=True, left_index=True, how="outer") # merging the two tables

# Joint Splice closure preparing the table
joint_1 = df6.loc[(df6['Status'] == 'Planned') & (df6['Enclosure Type'] == 'Joint')].groupby('Milestone').count()
joint_1 = joint_1.rename(columns={'Enclosure Type': 'Joint BOM'})
joint_1 = joint_1.drop(columns=['Status'])
joint_2 = df6.loc[(df6['Status'] == 'Spliced') & (df6['Enclosure Type'] == 'Joint')].groupby('Milestone').count()
joint_2 = joint_2.rename(columns={'Enclosure Type': 'Joint COM'})
joint_2 = joint_2.drop(columns=['Status'])
total_joint = pd.merge(joint_1, joint_2, right_index=True, left_index=True, how="outer")

# Merging the tables into one table to create the first table per Milestone
df_first = pd.merge(total_tree, total_poles, right_index=True, left_index=True, how="outer")
df_second = pd.merge(df_first, total_fibre_duct, right_index=True, left_index=True, how="outer")
df_third = pd.merge(df_second, total_chambers, right_index=True, left_index=True, how="outer")
df_forth = pd.merge(df_third, total_fibre_oh, right_index=True, left_index=True, how="outer")
df_fifth = pd.merge(df_forth, total_fibre_ug, right_index=True, left_index=True, how="outer")
df_sixth = pd.merge(df_fifth, total_fibre_oh_ug, right_index=True, left_index=True, how="outer")
df_seventh = pd.merge(df_sixth, total_odp, right_index=True, left_index=True, how="outer")
df_eigth = pd.merge(df_seventh, total_joint, right_index=True, left_index=True, how="outer")

# Calculating the total values of the first table individually
df_eigth.at['Grand Total', 'TC BOM'] = df_eigth['TC BOM'].sum()
df_eigth.at['Grand Total', 'TC COM'] = df_eigth['TC COM'].sum()
df_eigth.at['Grand Total', 'Pole BOM'] = df_eigth['Pole BOM'].sum()
df_eigth.at['Grand Total', 'Pole COM'] = df_eigth['Pole COM'].sum()
df_eigth.at['Grand Total', 'Civils BOM'] = df_eigth['Civils BOM'].sum()
df_eigth.at['Grand Total', 'Civils COM'] = df_eigth['Civils COM'].sum()
df_eigth.at['Grand Total', 'Chambers BOM'] = df_eigth['Chambers BOM'].sum()
df_eigth.at['Grand Total', 'Chambers COM'] = df_eigth['Chambers COM'].sum()
df_eigth.at['Grand Total', 'O/H BOM'] = df_eigth['O/H BOM'].sum()
df_eigth.at['Grand Total', 'O/H COM'] = df_eigth['O/H COM'].sum()
df_eigth.at['Grand Total', 'U/G BOM'] = df_eigth['U/G BOM'].sum()
df_eigth.at['Grand Total', 'U/G COM'] = df_eigth['U/G COM'].sum()
df_eigth.at['Grand Total', 'UG/OH BOM'] = df_eigth['UG/OH BOM'].sum()
df_eigth.at['Grand Total', 'UG/OH COM'] = df_eigth['UG/OH COM'].sum()
df_eigth.at['Grand Total', 'ODP BOM'] = df_eigth['ODP BOM'].sum()
df_eigth.at['Grand Total', 'ODP COM'] = df_eigth['ODP COM'].sum()
df_eigth.at['Grand Total', 'Joint BOM'] = df_eigth['Joint BOM'].sum()
df_eigth.at['Grand Total', 'Joint COM'] = df_eigth['Joint COM'].sum()

# Rename all specific rows in the databases for df1
df1.loc[df1['Milestone'] == 0, 'Milestone'] = 'MAH_0_000'

# Adding a new column to convert the Milestone head
df1['New Milestone'] = df1['Milestone'].str[4:5]
df2['New Milestone'] = df2['Milestone'].str[4:5]
df3['New Milestone'] = df3['Milestone'].str[4:5]
df4['New Milestone'] = df4['Milestone'].str[4:5]
df5['New Milestone'] = df5['Milestone'].str[4:5]
df6['New Milestone'] = df6['Milestone'].str[4:5]

# Creating a list to add the totals of df1
tc1 = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'N', '0']
ribbon1 = df1['New Milestone'].isin(tc1)

# TC BOM Tree Column
bom_trees = df1.loc[ribbon1].groupby(['New Milestone']).count()
bom_trees = bom_trees.drop(columns=['Milestone', 'Status'])
bom_trees = bom_trees.rename(columns={'tree_lengt': 'TC BOM'})

# TC COM Tree Column
com_trees = df1.loc[df1['Status'] == 'Completed', ['tree_lengt', 'New Milestone']].groupby('New Milestone').sum()
com_trees = com_trees.rename(columns={'tree_lengt': 'TC COM'})
total_com_trees_com_trees = pd.merge(bom_trees, com_trees, right_index=True, left_index=True, how="outer")

# BOM Poles Column
bom_poles = df2.loc[df2['Status'] == 'Planned', ['Status', 'New Milestone']].groupby('New Milestone').count()
bom_poles = bom_poles.rename(columns={'Status': 'Pole BOM'})

# COM Poles Column
com_pole = df2.loc[df2['Status'] == 'Built', ['Status', 'New Milestone']].groupby('New Milestone').count()
com_pole = com_pole.rename(columns={'Status': 'Pole COM'})
total_bom_com_poles = pd.merge(bom_poles, com_pole, right_index=True, left_index=True, how="outer")
total_bom_com_poles = total_bom_com_poles.rename(columns={'New Milestone': 'Ribbon'})

# Civils BOM Fibre Duct Column
bom_civils = df3.loc[df3['Status'] == 'Planned', ['Shape__Length', 'New Milestone']].groupby('New Milestone').sum()
bom_civils = bom_civils.rename(columns={'Shape__Length': 'Civils BOM'})

# Civils COM Fibre Duct Column
com_civils = df3.loc[df3['Status'] == 'Built', ['Status', 'New Milestone']].groupby('New Milestone').count()
com_civils = com_civils.rename(columns={'Status': 'Civils COM'})
total_bom_civils_com_civils = pd.merge(bom_civils, com_civils, right_index=True, left_index=True, how="outer")

# Chambers BOM Column
bom_chambers = df4.loc[df4['Status'] == 'Planned', ['Status', 'New Milestone']].groupby('New Milestone').count()
bom_chambers = bom_chambers.rename(columns={'Status': 'Chambers BOM'})

# Chambers COM Column
com_chambers = df4.loc[df4['Status'] == 'Built', ['Status', 'New Milestone']].groupby('New Milestone').count()
com_chambers = com_chambers.rename(columns={'Status': 'Chambers COM'})
total_bom_chambers_com_chambers = pd.merge(bom_chambers, com_chambers, right_index=True, left_index=True, how="outer")

# OH BOM Column Fibre Cable
bom_oh = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'OH')].groupby('New Milestone').sum()
bom_oh = bom_oh.rename(columns={'Shape__Length': 'O/H BOM'})

# OH COM Column Fibre Cable
com_oh = df5.loc[(df5['Placement'] == 'OH') & (df5['Status'] == 'Built'), ['Shape__Length', 'New Milestone']].groupby('New Milestone').sum()
com_oh = com_oh.rename(columns={'Shape__Length': 'O/H COM'})
total_bom_oh_com_oh = pd.merge(bom_oh, com_oh, right_index=True, left_index=True, how="left")


# UG BOM Column Fibre cable
bom_ug = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'UG')].groupby('New Milestone').sum()
bom_ug = bom_ug.rename(columns={'Shape__Length': 'U/G BOM'})

# UG COM Column Fibre cable
com_ug = df5.loc[(df5['Placement'] == 'UG') & (df5['Status'] == 'Built'), ['Shape__Length', 'New Milestone']].groupby('New Milestone').sum()
com_ug = com_ug.rename(columns={'Shape__Length': 'U/G COM'})
total_bom_ug_com_ug = pd.merge(bom_ug, com_ug, right_index=True, left_index=True, how="left")

# UG/OH BOM Column Fibre Cable
bom_ug_oh = df5.loc[(df5['Status'] == 'Planned') & (df5['Placement'] == 'UG/OH')].groupby('New Milestone').sum()
bom_ug_oh = bom_ug_oh.rename(columns={'Shape__Length': 'UG/OH BOM'})

# UG/OH COM Column Fibre Cable
com_ug_oh = df5.loc[(df5['Placement'] == 'UG/OH') & (df5['Status'] == 'Built'), ['Shape__Length', 'New Milestone']].groupby('New Milestone').sum()
com_ug_oh = com_ug_oh.rename(columns={'Shape__Length': 'UG/OH COM'})
total_bom_ug_oh_com_ug_oh = pd.merge(bom_ug_oh, com_ug_oh, right_index=True, left_index=True, how="outer")

# ODP BOM Column Splice Closures
bom_filt_odp = (df6['Status'] == 'Planned') & (df6['Enclosure Type'].isin(ODP)) # Creating a filter to add the ODP
bom_odp = df6.loc[bom_filt_odp].groupby('New Milestone').count() # Adding the filter and groupby Milestone
bom_odp = bom_odp.rename(columns={'Enclosure Type': 'ODP BOM'}) # Rename columns
bom_odp = bom_odp.drop(columns=['Status', 'Milestone']) # Drop columns

# ODP COM Column Splice Closures
com_filt_odp = (df6['Status'] == 'Spliced') & (df6['Enclosure Type'].isin(ODP2)) # Creating a filter to add the ODP
com_odp = df6.loc[com_filt_odp].groupby('New Milestone').count() # Adding the filter and groupby Milestone
com_odp = com_odp .rename(columns={'Enclosure Type': 'ODP COM'}) # Rename columns
com_odp = com_odp .drop(columns=['Status', 'Milestone']) # Drop columns
total_bom_odp_com_odp = pd.merge(bom_odp, com_odp, right_index=True, left_index=True, how="outer") # merging the two tables

# JOINT BOM Splice Closures
bom_joint = df6.loc[(df6['Status'] == 'Planned') & (df6['Enclosure Type'] == 'Joint')].groupby('New Milestone').count()
bom_joint = bom_joint.rename(columns={'Enclosure Type': 'Joint BOM'})
bom_joint = bom_joint.drop(columns=['Status', 'Milestone'])

# JOINT COM Splice Closure
com_joint = df6.loc[(df6['Status'] == 'Spliced') & (df6['Enclosure Type'] == 'Joint')].groupby('New Milestone').count()
com_joint = com_joint .rename(columns={'Enclosure Type': 'Joint COM'})
com_joint = com_joint .drop(columns=['Status', 'Milestone'])
total_bom_joint_com_joint = pd.merge(bom_joint, com_joint, right_index=True, left_index=True, how="outer")

# Merging the tables into one table to create the second table with totals per Ribbon
df_1 = pd.merge(total_com_trees_com_trees, total_bom_com_poles, right_index=True, left_index=True, how="outer")
df_2 = pd.merge(df_1, total_bom_civils_com_civils, right_index=True, left_index=True, how="outer")
df_3 = pd.merge(df_2, total_bom_chambers_com_chambers, right_index=True, left_index=True, how="outer")
df_4 = pd.merge(df_3, total_bom_oh_com_oh, right_index=True, left_index=True, how="outer")
df_5 = pd.merge(df_4, total_bom_ug_com_ug, right_index=True, left_index=True, how="outer")
df_6 = pd.merge(df_5, total_bom_ug_oh_com_ug_oh, right_index=True, left_index=True, how="outer")
df_7 = pd.merge(df_6, total_bom_odp_com_odp, right_index=True, left_index=True, how="outer")
df_8 = pd.merge(df_7, total_bom_joint_com_joint, right_index=True, left_index=True, how="outer").reset_index()
df_8 = df_8.rename(columns={'New Milestone': 'Ribbon'})


# Calculating the total values of the second table individually
df_8.at['Grand Total', 'TC BOM'] = df_8['TC BOM'].sum()
df_8.at['Grand Total', 'TC COM'] = df_8['TC COM'].sum()
df_8.at['Grand Total', 'Pole BOM'] = df_8['Pole BOM'].sum()
df_8.at['Grand Total', 'Pole COM'] = df_8['Pole COM'].sum()
df_8.at['Grand Total', 'Civils BOM'] = df_8['Civils BOM'].sum()
df_8.at['Grand Total', 'Civils COM'] = df_8['Civils COM'].sum()
df_8.at['Grand Total', 'Chambers BOM'] = df_8['Chambers BOM'].sum()
df_8.at['Grand Total', 'Chambers COM'] = df_8['Chambers COM'].sum()
df_8.at['Grand Total', 'O/H BOM'] = df_8['O/H BOM'].sum()
df_8.at['Grand Total', 'O/H COM'] = df_8['O/H COM'].sum()
df_8.at['Grand Total', 'U/G BOM'] = df_8['U/G BOM'].sum()
df_8.at['Grand Total', 'U/G COM'] = df_8['U/G COM'].sum()
df_8.at['Grand Total', 'UG/OH BOM'] = df_8['UG/OH BOM'].sum()
df_8.at['Grand Total', 'UG/OH COM'] = df_8['UG/OH COM'].sum()
df_8.at['Grand Total', 'ODP BOM'] = df_8['ODP BOM'].sum()
df_8.at['Grand Total', 'ODP COM'] = df_8['ODP COM'].sum()
df_8.at['Grand Total', 'Joint BOM'] = df_8['Joint BOM'].sum()
df_8.at['Grand Total', 'Joint COM'] = df_8['Joint COM'].sum()

# Transfer the Database to Excel writer to export the dataframe
writer = pd.ExcelWriter('Maghera Build Report.xlsx', engine='xlsxwriter')

# Adding a variable
sheet_name = 'Maghera Build Report'

# Setting the rows and the columns
df_eigth.to_excel(writer, sheet_name=sheet_name, startcol=1)
df_8.to_excel(writer, sheet_name=sheet_name, startcol=21)

# Adding a workbook to alter the tables
workbook = writer.book
worksheet = writer.sheets[sheet_name]
bold = workbook.add_format({'bold': True})

# Format the cells
cell_format = workbook.add_format({'bold': True, 'font_size': '14',
                                   'font_name': 'Calibri Light',
                                   'valign': 'vcenter', 'text_wrap': True})

#Adding the Milestone to Milestone columen header
worksheet.write_string(0, 1, 'Milestone', cell_format)

# Adding a bigger size for header
worksheet.set_row(0, 30)

# Adding a bigger size for footer
worksheet.set_row(33, 18)

# Adding a color format
format1 = workbook.add_format({'bg_color': '#DCE6F1', 'font_color': '#000000'}) # Light blue
format2 = workbook.add_format({'bg_color': '#D5E2B8', 'font_color': '#000000'}) # Light olive
format3 = workbook.add_format({'bg_color': '#ffffff', 'font_color': '#000000'}) # Light grey

# Declare columns with color light blue
worksheet.conditional_format('B1:T1', {'type': 'unique', 'format': format1})
worksheet.conditional_format('B32:T32', {'type': 'unique', 'format': format1})
worksheet.conditional_format('V1:AO1', {'type': 'unique', 'format': format1})
worksheet.conditional_format('V10:AO10', {'type': 'unique', 'format': format1})

# Set a list for the widths
colwidths = {}

# Store the defaults.
for col in range(50000):
    colwidths[col] = 15

# Calculate the width manually
colwidths[0] = 5
colwidths[1] = 17
colwidths[2] = 10
colwidths[3] = 10
colwidths[4] = 10
colwidths[5] = 10
colwidths[6] = 10
colwidths[7] = 15
colwidths[8] = 15
colwidths[9] = 15
colwidths[10] = 13
colwidths[11] = 13
colwidths[12] = 13
colwidths[13] = 10
colwidths[14] = 12
colwidths[15] = 12
colwidths[16] = 10
colwidths[17] = 10
colwidths[18] = 10
colwidths[19] = 10

# Then set the column widths.
for col_num, width in colwidths.items():
    worksheet.set_column(col_num, col_num, width)

# Saving to the local folder
writer.save()
