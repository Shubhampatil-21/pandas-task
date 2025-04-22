import pandas as pd
file = 'PandasTask.xlsx'

sheetOne = pd.read_excel(file, sheet_name=0)
sheetTwo = pd.read_excel(file, sheet_name=1)
tempSheetOne = sheetOne.drop_duplicates(subset="Feed Name")
tempSheetTwo = sheetTwo.drop_duplicates(subset="Feed Name")

mergeSheetOneandTwo = pd.merge(tempSheetOne, tempSheetTwo, on = ["Client", "Feed Name"], how="inner")

temp = mergeSheetOneandTwo.copy()
temp["Addition"] = temp.apply(lambda row: row["Import Product Count Today"] + row["Export Product Count Today"], axis = 1)
temp["Subtraction"] = temp.apply(lambda row: row["Import Product Count Today"] - row["Export Product Count Today"], axis = 1)
temp["Multiplication"] = temp.apply(lambda row: row["Import Product Count Today"] * row["Export Product Count Today"], axis = 1)
temp["Division"] = temp.apply(lambda row: row["Import Product Count Today"] / row["Export Product Count Today"] 
                              if row["Export Product Count Today"] != 0 else " ", 
                              axis = 1)

temp = temp.sort_values(by="Export Product Count Today", ascending=False)

with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    temp.to_excel(writer, sheet_name='Sheet3', index=False)