import pandas as pd
import numpy as np
import xlsxwriter

df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\STR_MAH_Full_OLT_CLD - Contractor\Maghera Vegetation Issues.xlsx')

# Convert the tree_lengt from string to numeric
df1['tree_lengt'] = pd.to_numeric(df1['tree_lengt'], errors='coerce')

# Filling the NaN values to 0
df1 = df1.fillna(0)

# Creating a lambda function calculating the total shape divided by count
func = lambda x: 100*x.count()/df1.shape[0]

# Adding pivot table to sum up all the tree per status and milestone
pivot1 = pd.pivot_table(df1, index=['Milestone'], values=['tree_lengt'], columns=['Status'],
                            aggfunc=np.sum, margins=True, margins_name='Grand Total')
print(pivot1)

# Adding a pivot table to allocate the percentage per status and milestone
pivot2 = pd.pivot_table(df1, index=['Milestone'], values=['tree_lengt'], columns=['Status'],
                            aggfunc=func, margins=True, margins_name='Grand Total')
print(pivot2)

# Transfer the Database to Excel writer to export the dataframe
writer = pd.ExcelWriter('Maghera Vegetation Progress.xlsx', engine='xlsxwriter')

# Adding variable
sheet_name = 'statistics'

# Setting the rows and the columns
pivot1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0,)
pivot2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8)

# Adding a woorkbook to alter the tables
workbook = writer.book
worksheet = writer.sheets[sheet_name]
bold = workbook.add_format({'bold': True})

# Format the cells
cell_format = workbook.add_format({'bold': True, 'font_size': '14',
                                   'font_name': 'Calibri Light',
                                   'valign': 'vcenter', 'text_wrap': True,})

# Adding a bigger size for header
worksheet.set_row(0, 40)

# Renaming the columns
worksheet.write('B1:E1', 'Summarizing the Length of the Trees per Status and Milestone', cell_format)
worksheet.write('J1:N1', 'Calculating the Total Percentage % of Status and Milestone', cell_format)

# Adding a color format
format1 = workbook.add_format({'bg_color': '#DCE6F1', 'font_color': '#000000'}) # Light blue
format2 = workbook.add_format({'bg_color': '#ffe6cc', 'font_color': '#000000'}) # Light orange
format3 = workbook.add_format({'bg_color': '#FF4500', 'font_color': '#000000'}) # Red
format4 = workbook.add_format({'bg_color': '#ff99ff', 'font_color': '#000000'}) # Purple
format5 = workbook.add_format({'bg_color': '#6CC417', 'font_color': '#000000'}) # Green
format6 = workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000'}) # Yellow
format7 = workbook.add_format({'bg_color': '#ffa600', 'font_color': '#000000'}) # Orange

# Adding appearing format for percentage
bor_format = workbook.add_format({'border': 2})
per_format = workbook.add_format({'num_format': '0.00%'})
comma_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)'})

# Declare columns with color light orange
worksheet.conditional_format('A2:A35', {'type': 'unique', 'format': format2})
worksheet.conditional_format('I2:I35', {'type': 'unique', 'format': format2})

# Declare the columns based on pie chart Sum
worksheet.conditional_format('B2:B2', {'type': 'unique', 'format': format3})
worksheet.conditional_format('B35:B35', {'type': 'unique', 'format': format3})
worksheet.conditional_format('C2:C2', {'type': 'unique', 'format': format4})
worksheet.conditional_format('C35:C35', {'type': 'unique', 'format': format4})
worksheet.conditional_format('D2:D2', {'type': 'unique', 'format': format5})
worksheet.conditional_format('D35:D35', {'type': 'unique', 'format': format5})
worksheet.conditional_format('E2:E2', {'type': 'unique', 'format': format6})
worksheet.conditional_format('E35:E35', {'type': 'unique', 'format': format6})
worksheet.conditional_format('F2:F2', {'type': 'unique', 'format': format7})
worksheet.conditional_format('F35:F35', {'type': 'unique', 'format': format7})

# Declare the columns based on pie chart percentage
worksheet.conditional_format('J2:J2', {'type': 'unique', 'format': format3})
worksheet.conditional_format('J35:J35', {'type': 'unique', 'format': format3})
worksheet.conditional_format('K2:K2', {'type': 'unique', 'format': format4})
worksheet.conditional_format('K35:K35', {'type': 'unique', 'format': format4})
worksheet.conditional_format('L2:L2', {'type': 'unique', 'format': format5})
worksheet.conditional_format('L35:L35', {'type': 'unique', 'format': format5})
worksheet.conditional_format('M2:M2', {'type': 'unique', 'format': format6})
worksheet.conditional_format('M35:M35', {'type': 'unique', 'format': format6})
worksheet.conditional_format('N2:N2', {'type': 'unique', 'format': format7})
worksheet.conditional_format('N35:N35', {'type': 'unique', 'format': format7})


# Adding the format for percentages pivot table
worksheet.conditional_format('J2:J35', {'type': 'unique', 'format': comma_format})
worksheet.conditional_format('K2:K35', {'type': 'unique', 'format': comma_format})
worksheet.conditional_format('L2:L35', {'type': 'unique', 'format': comma_format})
worksheet.conditional_format('M2:M35', {'type': 'unique', 'format': comma_format})
#worksheet.conditional_format('N4:N35', {'type': 'unique', 'format': comma_format})

# Adding the chart type
chart = workbook.add_chart({'type': 'pie'})

# Adding the Series, Name, Columns, Values, style,
chart.add_series({
    'name': 'Total Percentage % per Status',
    'categories': '=statistics!$J$2:$M$2',
    'values': '=statistics!$J$35:$M$35',
    'data_labels': {'percentage': True, 'leader_lines': True,
                     'category': True, 'legend_key': True},
    'points': [
            {'fill': {'color': '#FF4500'}}, # Red
            {'fill': {'color': '#ff99ff'}}, # Purple
            {'fill': {'color': '#6CC417'}}, # Green
            {'fill': {'color': '#FFFF00'}}, # yellow
            {'fill': {'color': '#ffa600'}} # Orange
            ],
    })

# Adding the chart to excel sheet and increase the scale of the chart
worksheet.insert_chart('Q3', chart, {'x_scale': 1.3, 'y_scale': 1.4})

# Set a list for the widths
colwidths = {}

# Store the defaults.
for col in range(50000):
    colwidths[col] = 15

# Calculate the width manually
colwidths[0] = 12
colwidths[1] = 9
colwidths[2] = 12
colwidths[3] = 12
colwidths[4] = 13
colwidths[5] = 13
colwidths[6] = 18
colwidths[7] = 6
colwidths[8] = 15
colwidths[9] = 13
colwidths[10] = 13
colwidths[11] = 13
colwidths[12] = 13
colwidths[13] = 13
colwidths[14] = 13

# Then set the column widths.
for col_num, width in colwidths.items():
    worksheet.set_column(col_num, col_num, width)

writer.save()
