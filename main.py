import openpyxl
import pandas as pd

LIST_FILE = "./resources/input/list.xlsx"
TEMPLATE_FILE = "./resources/input/template.xlsx"

# エクセルファイルを読み込む
df_list = pd.read_excel(LIST_FILE, sheet_name="Sheet1", header=8, index_col=1)

# 必要な列のみ抽出する
target_data = df_list[['施設名称', '台数計算定員÷5']]

for index,row in target_data.iterrows():
    # 新規ファイル作成
    book = openpyxl.load_workbook(filename=TEMPLATE_FILE)
    # シートの1枚目を指定する
    sheet = book.worksheets[0]
    # C3のセルに施設名称を設定する
    sheet['C4'] = row['施設名称']
    sheet['D11'] = '： ' + row['施設名称']
    sheet['G19'] = row['台数計算定員÷5']

    # outputディレクトリ配下にtest_インデックス名.xlsxという名前で保存する
    book.save('./resources/output/' + str(index) + row['施設名称'] +  '.xlsx')
    book.close()
