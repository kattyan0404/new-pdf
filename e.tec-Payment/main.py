import tabula
import re
import openpyxl

def convert_int(num):
    num = re.sub("\¥|\,","",num)
    num = int(num)
    return num

def convert_1d_to_2d(l, cols):
    return [l[i:i + cols] for i in range(0, len(l), cols)]


excel_path = 'e.tec-Payment\請求書 模式データ 空ファイル.xlsx'


path = "F:\python_code\e.tec-Payment\B601アーステック様.pdf"
tables = tabula.read_pdf(path, pages='all', lattice=True, pandas_options={'header': None})

# tables[ページ数] [列] [行]

# 税率　固定変数
tax = float(0.1)

print()
x_length = len(tables[2][0])

# tabulaからのデータを1次元リストに変換
tables_list = []
for y in range(0, x_length):
    for x in range(0, 6):
        tables_list.append(tables[2][x][y])

# 1次元リストを2次元リストに変換
list_length = len(tables_list)
tables_list = convert_1d_to_2d(tables_list, 6)
tables_list_length = len(tables_list)

    

# 金額の有無で判定する
# 金額の値があった場合はdt_listに代入
dt_list = []

# 作業用リストにデータ追加
for i in range(tables_list_length):
    if tables_list[i][5] != '¥0' and type(tables_list[i][5]) != type(float(0.0)):
            dt_list.append(tables_list[i][5])

dt_list_length = len(dt_list)


# dt_listの個数分dt_listの配列要素をdts_listに代入
dts_list = []
for i in range(tables_list_length):
    for x in range(dt_list_length):
        if tables_list[i][5] == dt_list[x]:
            dts_list.append(tables_list[i])
dts_list_length = len(dts_list)
for i in range(dts_list_length):
    print(dts_list[i][1])
    if dts_list[i][1].startswith('宿泊費'):
        dts_list[i], dts_list[i+1] = dts_list[i+1], dts_list[i]
        break


print(dts_list)

dts_hundring = tables_list_length - 11
# print(tables_list[dts_hundring])

# 金額参照で抜けている部分を手動で補う判定式
if tables_list[dts_hundring][3] != '¥0' and type(tables_list[dts_hundring][3]) != type(float(0.0)):
    dts_list.append(tables_list[dts_hundring])
if tables_list[dts_hundring + 1][3] != '¥0' and type(tables_list[dts_hundring][3]) != type(float(0.0)):
    dts_list.append(tables_list[dts_hundring + 1])
dts_list_length = len(dts_list)
dts_data_length = len(dts_list[0])




# エクセル用データに変換

# 項目名判定・変換
for i in range(dts_list_length):
    if dts_list[i][1] == '宿泊費(実費)4/28~5/4':
        dts_list[i][1] = '宿泊料金'
    if dts_list[i][1] == '出張手当':
        dts_list[i][1] = '出張割増'
    if dts_list[i][1] == '旅費(JR・タクシー・フェリー)':
        jr = input('jrの金額を入力してください : ')
        taxi = input('タクシーの金額を入力してください : ')
        ferry = input('フェリーの金額を入力してください : ')

print(f'jr料金 : {jr}円')
print(f'タクシー料金 : {taxi}円')

# 単位判定変換
for i in range(dts_list_length):
    if dts_list[i][3] == '時間':
        dts_float_conver = float(dts_list[i][2])
        dts_int_conver = int(dts_float_conver)
        dts_list[i][2] = str(dts_int_conver) + 'H'
    if dts_list[i][3] == '件':
        dts_list[i][2] = str(dts_int_conver) + '日'

print(dts_list)


# job-Nomber
job_No = tables[0][1][1]
print(job_No)


# total-amount
total_amount = tables[0][1][2]
print(total_amount)


total_amount = convert_int(total_amount)

tax_amount = total_amount * tax
tax_total_amount = tax_amount + total_amount
tax_total_amount = int(tax_total_amount)
print(tax_total_amount)




print()
print('書き込み開始')
# ブックを取得
wb = openpyxl.load_workbook(excel_path)

# シートを取得
sb = wb['Sheet1']


# ヘッダー部分の書き込み
sb['C9'] = job_No
sb['D14'] = tax_total_amount
sb['K14'] = tax_amount

# 請求内容の書き込み
cell_len = 15
for y in range(dts_list_length):
    cell_len += 2
    for x in range(dts_data_length):
        sb[f'B{cell_len}'] = dts_list[y][1]
        sb[f'G{cell_len}'] = dts_list[y][2]
        sb[f'H{cell_len}'] = dts_list[y][4]
        sb[f'I{cell_len}'] = dts_list[y][5]

wb.save(excel_path)

print('書き込み終了')