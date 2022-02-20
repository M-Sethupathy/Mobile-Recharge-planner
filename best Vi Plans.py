# pip install pandas
# pip install xlsxwriter

import requests as req
import json
import pandas as pd
from datetime import datetime as dati

planURL = "https://www.myvi.in/content/dam/VIL/HARMONIZED/PREPAID/PACK/MH"
data = req.get(planURL)
if data.status_code != 200:
    print("Error getting plans info")
    exit()
else :
    print("data loaded successfully")

planData = json.loads(data.content)["DATA"]

neededKeysMiniExcel = [
    "COMBO_TYPE_ATTR",
    "DATA_LINE_1",
    "PRODUCT-NAME",
    "READ_MORE",
    "RECHARGENAME_ATTR",
    "RECHARGE_SUBTYPE",
    "SMS_LINE_1",
    "UnitCost",
    "VALIDITY_ATTR",
]
def getPlanInfoMiniExcel(category, plan):
    viPlansJson["category_name"].append(category)
    if plan["VALIDITY_ATTR"] == "0":
        ratePerDay = float(plan["UnitCost"])
    else:
        ratePerDay = float(plan["UnitCost"]) / float(plan["VALIDITY_ATTR"])
    
    viPlansJson["RsperDay"].append(ratePerDay)
    for i in neededKeysMiniExcel:
        val = ""
        if i in plan and plan[i]:
            val = plan[i]
        viPlansJson[i].append(val)

viPlansJson = {}

viPlansJson["category_name"] =  []
viPlansJson["RsperDay"] = []
viPlansJson["UnitCost"] = []
viPlansJson["VALIDITY_ATTR"] = []
viPlansJson["DATA_LINE_1"] = []
viPlansJson["PRODUCT-NAME"] =  []
viPlansJson["RECHARGE_SUBTYPE"] =  []
viPlansJson["RECHARGENAME_ATTR"] =  []
viPlansJson["COMBO_TYPE_ATTR"] =  []
viPlansJson["SMS_LINE_1"] =  []
viPlansJson["READ_MORE"] =  []

for i in planData:
    for j in i["subcategorylist"]:
        if "subcategorylist" not in j:
            getPlanInfoMiniExcel(i["category_name"], j)

VIP = pd.DataFrame(viPlansJson)
VIP.sort_values("RsperDay", inplace=True, ignore_index=True)
xlFileName = "ViPlans - {} - auto aligned.xlsx".format(dati.now().strftime("%Y-%m-%d_%H-%M-%S"))
writer = pd.ExcelWriter(xlFileName, engine ='xlsxwriter')

VIP.to_excel(writer, sheet_name="Sheet1")
# VIP['READ_MORE'].str.wrap(100) #to set max line width of 100
worksheet = writer.sheets["Sheet1"]  # pull worksheet object
for idx, col in enumerate(VIP):  # loop through all columns
    series = VIP[col]
    if series.name != "READ_MORE":
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name)) * 2  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    else:
        max_len = max((
            0,  # len of largest item
            len(str(series.name)) * 2  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx+1, max_len)  # set column width
print("Excel file saved successfully", xlFileName)
writer.save()