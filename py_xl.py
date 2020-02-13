import openpyxl
import sys
import pprint

# pip install openpyxl

#第一引数をファイル名とする
args = sys.argv
excel_file = args[1]

#ワークブックの指定（Excelファイルの読み込み）
workbook = openpyxl.load_workbook(excel_file)

#シートの取得
sheet = workbook["Skill_Sheet"]

#指定セルの取得
cell = sheet['B5']

print(cell.value)

#セルへの書き込み
sheet['B5'] = '岡本　次郎'

#pprint.pprint(sheet['A26:E54'])

#配列に読込
ary_xl = sheet['A26:E54']

#検索したい文字列をキー入力
search_str = input('検索したい文字列をキー入力>>')

#入力した文字列を業務内容列[N][2]から探すi=0
for tmp_obj1 in ary_xl:
    for tmp_obj2 in tmp_obj1:
        tmp_text:str = str(tmp_obj2.value)
        if search_str in tmp_text:
            print( str(tmp_obj2.coordinate) + " に有った")
            #該当セルを表示する
            print(tmp_text)

workbook.save(excel_file)