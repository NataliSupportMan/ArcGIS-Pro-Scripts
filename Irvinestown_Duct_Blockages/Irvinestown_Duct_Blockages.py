import pandas as pd

df1 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Irvinestown_Fibrus\STR_IRV1_COL_BT_Duct_Blockage_Coordinates.xlsx')
df2 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Irvinestown_Fibrus\Irvinestown_Duct_Blockage_Online_Excel.xlsx')
df3 = pd.read_excel(r'C:\Users\NataliSuportman\Documents\ArcGIS\Projects\Irvinestown_Fibrus\STR_IRV1_COL_BT_Duct_Blockage_R_R_old_points.xlsx')
df2 = df2[['Editor']]
df3 = df3[['Value in both lists', 'point_x old']]

# Concat all the tables to 1 table
all_table = pd.concat([df1, df2, df3], axis=1)

# Creating a function and compare the two columns
def compare(df):
    if df['POINT_X'] == df['point_x old']:
        return 'Exist'
    elif df['POINT_X'] != df['point_x old']:
        return 'Not Exist'
    else:
        pass
all_table['Compare Values POINT_X/point_x old'] = all_table.apply(compare, axis=1)

# Re order the columns
all_table = all_table[['OID', 'GlobalID', 'Editor', 'Comments', 'Blockage Type', 'Pia Blockage',
                        'POINT_X', 'Compare Values POINT_X/point_x old', 'point_x old', 'POINT_Y', 'DDLat', 'DDLon']]
print(all_table.to_string())

# Export to Excel the new table
all_table.to_excel('Irvinestown Duct Blockages.xlsx', index=False)