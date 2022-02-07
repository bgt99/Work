
import pandas as pd
df = pd.read_csv('TempCol2022012513200154_SUM.csv')
writer = pd.ExcelWriter('TempCol2022012513200154_SUM.xlsx')
df.to_excel(writer, index=False)
writer.save()
