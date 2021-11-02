################################################## EXPORTING from Feature Layer to xlsx files ####################################################################

import arcpy

ribbon_layers = [
        ["STR GRS1 Full OLT LLD - Vegetation Issues", "STR GRS1 Full OLT LLD - Vegetation Issues.xlsx"],
        ["STR GRS1 Full OLT CLD - Poles", "STR GRS1 Full OLT CLD - Poles.xlsx"],
        ["STR GRS1 Full OLT CLD - Chambers", "STR GRS1 Full OLT CLD - Chambers.xlsx"],
        ["STR GRS1 Full OLT CLD - Fibre Duct", "STR GRS1 Full OLT CLD - Fibre Duct.xlsx"],
        ["STR GRS1 Full OLT CLD - Fibre Cable", "STR GRS1 Full OLT CLD - Fibre Cable.xlsx"],
        ["STR GRS1 Full OLT CLD - Splice Closures", "STR GRS1 Full OLT CLD - Splice Closures.xlsx"],
        ["STR GRS1 Full OLT CLD - Drop Wires", "STR GRS1 Full OLT CLD - Drop Wires.xlsx"],
        ["DFE - Garrison Contractor B", "DFE - Garrison Contractor B.xlsx"]
        ]

for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())