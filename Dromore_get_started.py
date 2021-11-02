######################################## CALCULATE FIELDS from Feature layers for all Ribbons in Dromore Layer #############################################

import arcpy

current_project = arcpy.mp.ArcGISProject("CURRENT")
maps = current_project.listMaps()

ribbon_list = [
    ['Ribbon A/STR_DRO1_RibbonA_COL_Duct For TRR', 'L0bt_duct_route', 'L0STR_DRO1_RibbonA_COL_Duct_For_TRR_KN'],
    ['Ribbon B/STR_DRO1_RibbonB_COL_Duct For TRR', 'L0bt_duct_route', 'L0STR_DRO1_RibbonB_COL_Duct_For_TRR_KN'],
    ['Ribbon C/STR_DRO1_RibbonC_COL_Duct For TRR', 'L0bt_duct_route', 'L0STR_DRO1_RibbonC_COL_Duct_For_TRR_KN']
    ]
for ribbon in ribbon_list:
    field_pairs = [[ribbon[1] + ".duct_stat", ribbon[2] + ".duct_stat"],
                   [ribbon[1] + ".duct_cap", ribbon[2] + ".duct_cap"],
                   [ribbon[1] + ".sub_inst", ribbon[2] + ".sub_inst"],
                   [ribbon[1] + ".de_silt", ribbon[2] + ".de_silt"],
                   [ribbon[1] + ".comments", ribbon[2] + ".comments"],
                   [ribbon[1] + ".status", ribbon[2] + ".status"]]
    fc = ribbon[0]
    for field_pair in field_pairs:
        online_field = field_pair[0]
        kn_field = field_pair[1]
        selection_fc = arcpy.management.SelectLayerByAttribute(fc, 'NEW_SELECTION', where_clause=online_field + ' <> ' + kn_field)
        arcpy.management.CalculateField(selection_fc, online_field, '!' + kn_field + '!', 'PYTHON3')
print(time.asctime())

for map_sel in maps:
    map_sel.clearSelection()

import arcpy

ribbon_layers = [
        ["Ribbon A/STR_DRO1_RibbonA_COL_Duct For TRR", "STR_DRO1_RibbonA_COL_Duct For TRR.xlsx"],
        ["Ribbon B/STR_DRO1_RibbonB_COL_Duct For TRR", "STR_DRO1_RibbonB_COL_Duct For TRR.xlsx"],
        ["Ribbon C/STR_DRO1_RibbonC_COL_Duct For TRR", "STR_DRO1_RibbonC_COL_Duct For TRR.xlsx"],
        ["Dromore Duct Blockage", "Dromore Duct Blockage.xlsx"],
        ]

for ribbon_layer in ribbon_layers:
    layer_name = ribbon_layer[0]
    export_files = ribbon_layer[1]
    arcpy.TableToExcel_conversion(layer_name, export_files, 'ALIAS', 'DESCRIPTION')
print(time.asctime())

