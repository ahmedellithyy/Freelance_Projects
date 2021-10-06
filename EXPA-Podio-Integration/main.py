from dataProcessing import *
from podioUpdating import *



# 0 index -> EXPA LC Code
# 1 index -> ICX App ID
# 2 index -> OGX App ID
# if -1 in App ID -> it means LC doesn't have app in this area
lc_codes = {
    "6th October":	[2820,-1,26580039],
    "AAST-Alex": [1788, 26580315, 26580312],
    "AAST-Cairo":	[1322,26580054,26580058],
    "Ain Shams":	[1789,26580192,26580196],
    "Alexandria":	[899,-1,26580208],
    "AUC":	[1489,26580129,26580126],
    "Beni Suef": [2126,-1,26580166],
    "Cairo University": [1064,26580073,26580070],
    "Damietta":	[109,26579999,26579996],
    "GUC":	[257,26580014,26580011],
    "Helwan":	[2124,26580082,26580086],
    "KSU":	[2524,-1,26580298],
    "Luxor":	[2114,26580240,26580244],
    "Mansoura":	[171,26580324,26580328],
    "Menofia":	[1727,26580254,26580258],
    "MIU":	[2125,26580287,26580284],
    "MSA":	[2817,-1,26580106],
    "MUST":	[2818,-1,26580342],
    "Suez":	[15,26580028,26580025],
    "Tanta": [1725,26580353,26580357],
    "Zagazig":[1114,26580268,26580272]

}
# write the date in this format dd/MM/yyyy
startdate = "01/08/2021"
enddate = "01/09/2021"

lc_codes_values = lc_codes.values()
lc_codes_names = list(lc_codes.keys())
resultSum = 0
i = 0
for lc in lc_codes_values:
    expa_code = lc[0]
    icx_app_id = lc[1]
    ogx_app_id = lc[2]
    result_OGX_APDs, result_OGX_REs, result_OGX_FIs, result_ICX_APDs, result_ICX_REs, result_ICX_FIs = extract_data(
        expa_code, startdate, enddate)
    print(result_OGX_APDs,"\n",result_OGX_REs,"\n",result_OGX_FIs)
    print(result_ICX_APDs,"\n",result_ICX_REs,"\n",result_ICX_FIs)
    resultSumOGX = len(result_OGX_APDs) + len(result_OGX_REs) + len(result_OGX_FIs)
    resultSumICX = len(result_ICX_APDs) + len(result_ICX_REs) + len(result_ICX_FIs)
    if icx_app_id != -1:
         newAddedItems, notAddedItems = podioUpdateAPDsICX(result_ICX_APDs, icx_app_id)
         printValues(result_ICX_APDs, newAddedItems, notAddedItems, "ICX APDs")
         newAddedItems, UpdatedItems = podioUpdateREsICX(result_ICX_REs, icx_app_id)
         printValues(result_ICX_REs, newAddedItems, UpdatedItems, "ICX REs")
         newAddedItems, UpdatedItems = podioUpdateREsICX(result_ICX_FIs, icx_app_id)
         printValues(result_ICX_REs, newAddedItems, UpdatedItems, "ICX FIs")
         print("ICX Finished Uploading")
         print(resultSumICX)
    newAddedItems = podioUpdateAPDsOGX(result_OGX_APDs, ogx_app_id)
    printValues(result_OGX_APDs, newAddedItems, 0, "OGX APDs")
    newAddedItems, UpdatedItems = podioUpdateREsOGX(result_OGX_REs, ogx_app_id)
    printValues(result_OGX_REs, newAddedItems, UpdatedItems, "OGX REs")
    newAddedItems, UpdatedItems = podioUpdateREsOGX(result_OGX_FIs, ogx_app_id)
    printValues(result_OGX_FIs, newAddedItems, UpdatedItems, "OGX FIs")
    print("OGX Finished Uploading")
    print(lc_codes_names[i])
    i = i + 1
    print(resultSumOGX)


