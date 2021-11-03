import arcpy

# Creating variables for the layers
in_features = "STR_LRN1_COL_BT Duct Blockage_R_R"
feature_path = "C:/Users/NataliSuportman/Documents/ArcGIS/Projects/Larne_Fibrus/Default.gdb"
out_features_name = "STR_LRN1_COL_BT Duct Blockage_R_R_New"


# Creating the Feature Class to Feature Class layer
arcpy.FeatureClassToFeatureClass_conversion(in_features,
                                            feature_path,
                                            out_features_name)
# Adding the point to selected layer
arcpy.AddXY_management(out_features_name)

# set parameter values to convert coordinate notation
out_features_name = "STR_LRN1_COL_BT Duct Blockage_R_R_New"
output_points = 'C:/Users/NataliSuportman/Documents/ArcGIS/Projects/Larne_Fibrus/Default.gdb/STR_LRN1_COL_BT_Duct_Blockage_Coordinates'
x_field = 'POINT_X'
y_field = 'POINT_Y'
input_format = 'SHAPE'
output_format = 'DD 2'

try:
    arcpy.ConvertCoordinateNotation_management(out_features_name, output_points, x_field, y_field,
                                               input_format, output_format)
    print(arcpy.GetMessages(0))

except arcpy.ExecuteError:
    print(arcpy.GetMessages(2))

# Exporting the selected layer
import_layer = 'STR_LRN1_COL_BT_Duct_Blockage_Coordinates'
export_layer = 'STR_LRN1_COL_BT_Duct_Blockage_Coordinates.xlsx'
arcpy.TableToExcel_conversion(import_layer, export_layer, 'ALIAS', 'DESCRIPTION')
print(time.asctime())