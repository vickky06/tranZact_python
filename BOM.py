import pandas as pd 
import openpyxl
import json
df = pd.read_excel("BOM.xlsx",engine = "openpyxl")
result = df.to_json(orient="records")
filterdata = df["Item Name"].unique() 
maxDict = {}
parsed = json.loads(result) 
for el in parsed:
    if el["Item Name"]:
        if  el["Item Name"] not in maxDict:
            maxDict[ el["Item Name"]] = el["Level"]
        else:
            maxDict[el["Item Name"]] =  max( float(el["Level"]),  float(maxDict[el["Item Name"]]))


for i in filterdata:
    if i in maxDict.keys():
        start+=1
        maxval = maxDict[i] 
        for j in range(int(maxval)):
            a = df[(df["Item Name"].str.contains(i,na=False)) & (df["Level"] == (j+1))] 
            a.to_excel(i+str(j+1)+"_subgroup.xlsx")


print("We are done ")
