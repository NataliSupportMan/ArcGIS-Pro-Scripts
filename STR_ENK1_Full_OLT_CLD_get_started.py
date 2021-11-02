
################################################## EXPORTING from Feature Layer to xlsx files ####################################################################

import arcpy

ribbon_layers = [
        ["STR ENK1 Full OLT LLD Vegetation Issues", "STR ENK1 Full OLT LLD Vegetation Issues.xlsx"],
        ["STR ENK1 Full OLT CLD - Poles", "STR ENK1 Full OLT CLD - Poles.xlsx"],
        ["STR ENK1 Full OLT CLD - Chambers", "STR ENK1 Full OLT CLD - Chambers.xlsx"],
        ["STR ENK1 Full OLT CLD - Fibre Duct", "STR ENK1 Full OLT CLD - Fibre Duct.xlsx"],
        ["STR ENK1 Full OLT CLD - Fibre Cable", "STR ENK1 Full OLT CLD - Fibre Cable.xlsx"],
        ["STR ENK1 Full OLT CLD - Splice Closures", "STR ENK1 Full OLT CLD - Splice Closures.xlsx"],
        ["STR ENK1 Full OLT CLD - Drop Wires", "STR ENK1 Full OLT CLD - Drop Wires.xlsx"],
        ["DFE - Enniskillen Contractor B", "DFE - Enniskillen Contractor B.xlsx"],
        ]
for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())

