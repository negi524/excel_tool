import openpyxl
import pandas as pd

LIST_FILE = "./resources/input/list.xlsx"
TEMPLATE_FILE = "./resources/input/template.xlsx"

KEY_1 = '施設名称'
KEY_2 = '台数計算'

# エクセルファイルを読み込む
df_list = pd.read_excel(LIST_FILE, header=8, index_col=1)

# 必要な列のみ抽出する
target_data = df_list[[KEY_1, KEY_2]]

for index,row in target_data.iterrows():
    # 新規ファイル作成
    book = openpyxl.load_workbook(filename=TEMPLATE_FILE)
    # シートの1枚目を指定する
    sheet = book.worksheets[0]
    # C3のセルに施設名称を設定する
    sheet['C4'] = row[KEY_1]
    sheet['D11'] = '： ' + row[KEY_1]
    sheet['G19'] = row[KEY_2]

    # outputディレクトリ配下にインデックス名+施設名.xlsxという名前で保存する
    book.save('./resources/output/' + str(index) + row[KEY_1] +  '.xlsx')
    book.close()
