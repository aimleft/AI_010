import pandas as pd
import openpyxl
import pandas as pd

df = pd.read_excel('df_p.xlsx', sheet_name="31E")

print(df.to_string())
print(df.head(1))

print(df.iloc[0, 0], )

book = openpyxl.load_workbook('df_p.xlsx')

worksheet = book["31E"]

ce = worksheet.cell(2, 3)
fill = ce.fill

cell_color = fill.start_color.rgb

print(cell_color)
if cell_color == "FFFF0000":
    print("RED")
elif cell_color == "FF00FF00":
    print("GREEN")