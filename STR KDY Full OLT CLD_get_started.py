################################################## EXPORTING from Feature Layer to xlsx files ####################################################################

import arcpy

ribbon_layers = [
        ["STR KDY Ribbon Full OLT LLD - Vegetation Issues", "STR KDY Ribbon Full OLT LLD - Vegetation Issues.xlsx"],
        ["STR KDY1 Full OLT CLD - Poles", "STR KDY1 Full OLT CLD - Poles.xlsx"],
        ["STR KDY1 Full OLT CLD - Chambers", "STR KDY1 Full OLT CLD - Chambers.xlsx"],
        ["STR KDY1 Full OLT CLD - Fibre Duct", "STR KDY1 Full OLT CLD - Fibre Duct.xlsx"],
        ["STR KDY1 Full OLT CLD - Fibre Cable", "STR KDY1 Full OLT CLD - Fibre Cable.xlsx"],
        ["STR KDY1 Full OLT CLD - Splice Closures", "STR KDY1 Full OLT CLD - Splice Closures.xlsx"],
        ["STR KDY1 Full OLT CLD - Drop Wires", "STR KDY1 Full OLT CLD - Drop Wires.xlsx"],
        ["DFE - Keady Contractor B", "DFE - Keady Contractor B.xlsx"],
        ]

for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())
