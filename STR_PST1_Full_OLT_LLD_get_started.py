
################################################## EXPORTING from Feature Layer to xlsx files ####################################################################

import arcpy

ribbon_layers = [
        ["STR PST1 Full OLT LLD Vegetation Issues", "STR PST1 Full OLT LLD Vegetation Issues.xlsx"],
        ["STR PST1 Full OLT CLD - Poles", "STR PST1 Full OLT CLD - Poles.xlsx"],
        ["STR PST1 Full OLT CLD - Chambers", "STR PST1 Full OLT CLD - Chambers.xlsx"],
        ["STR PST1 Full OLT CLD - Fibre Duct", "STR PST1 Full OLT CLD - Fibre Duct.xlsx"],
        ["STR PST1 Full OLT CLD - Fibre Cable", "STR PST1 Full OLT CLD - Fibre Cable.xlsx"],
        ["STR PST1 Full OLT CLD - Splice Closures", "STR PST1 Full OLT CLD - Splice Closures.xlsx"],
        ["STR PST1 Full OLT CLD - Drop Wires", "STR PST1 Full OLT CLD - Drop Wires.xlsx"],
        ["DFE - Portstewart Contractor B", "DFE - Portstewart Contractor B.xlsx"],
        ]

for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())