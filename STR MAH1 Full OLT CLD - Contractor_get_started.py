
################################################## EXPORTING from Feature Layer to xlsx files ####################################################################

import arcpy

ribbon_layers = [
        ["Maghera Vegetation Issues", "Maghera Vegetation Issues.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Poles", "STR MAH1 Full OLT CLD - Contractor - Poles.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Chambers", "STR MAH1 Full OLT CLD - Contractor - Chambers.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Fibre Duct", "STR MAH1 Full OLT CLD - Contractor - Fibre Duct.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Fibre Cable", "STR MAH1 Full OLT CLD - Contractor - Fibre Cable.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Splice Closures", "STR MAH1 Full OLT CLD - Contractor - Splice Closures.xlsx"],
        ["STR MAH1 Full OLT CLD - Contractor - Drop Wires", "STR MAH1 Full OLT CLD - Contractor - Drop Wires.xlsx"],
        ["DFE - Maghera Contractor B", "DFE - Maghera Contractor B.xlsx"],
        ]

for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())



