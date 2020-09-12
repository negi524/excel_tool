import openpyxl
import pandas as pd

LIST_FILE = "./resources/list.xlsx"

# エクセルファイルを読み込む
df_list = pd.read_excel(LIST_FILE, sheet_name="Sheet1", header=8, index_col=1)

# 必要な列のみ抽出する
data = df_list.loc[:,['施設名称', '台数計算定員÷5']]
print(data)