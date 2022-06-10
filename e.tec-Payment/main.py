import tabula
import re
import openpyxl
import math

# tables[ページ数] [列] [行]

def convert_int(num):
    num = re.sub("\¥|\,","",num)
    num = int(num)
    return num


def convert_1d_to_2d(l, cols):
    return [l[i:i + cols] for i in range(0, len(l), cols)]
print('-----------------------------------------------------------------------')

print('請求書作成処理　開始')

print('-----------------------------------------------------------------------')


excel_path = 'e.tec-Payment\請求書 模式データ 空ファイル.xlsx'


path = "F:\python_code\e.tec-Payment\B601アーステック様.pdf"
tables = tabula.read_pdf(path, pages='all', lattice=True, pandas_options={'header': None})


tax = float(0.1)  # 税率　固定変数


str_cell_data = [1, '式']  # セルに表示する文字列


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



dts_hundring = tables_list_length - 11
# print(tables_list[dts_hundring])

# 金額参照で抜けている部分を手動で補う判定式
for i in range(0,2):
    if tables_list[dts_hundring + i][3] != '¥0' and type(tables_list[dts_hundring][3]) != type(float(0.0)):
        dts_list.append(tables_list[dts_hundring + i])

dts_list_length = len(dts_list)




# エクセル用データに変換

# 項目名判定・変換
for i in range(dts_list_length):
    if dts_list[i][1].startswith('宿泊費'):
        dts_list[i][1] = '宿泊料金'
        dts_list[i], dts_list[i+1] = dts_list[i+1], dts_list[i]
        dts_list[i+1][3] = '実費'


    if dts_list[i][1] == '出張手当':
        dts_list[i][1] = '出張割増'
    if dts_list[i][1] == '旅費(JR・タクシー・フェリー)':
        dts_list[i][3], dts_list[i][5] = dts_list[i][5], dts_list[i][3]
        dts_list[i][2], dts_list[i][3] = dts_list[i][3], dts_list[i][2]
        jr = input('jrの金額を入力してください : ')
        taxi = input('タクシーの金額を入力してください : ')
        ferry = input('フェリーの金額を入力してください : ')

public_expenses_len = 0
public_expense_vehicles = ['JR料金', 'タクシー代', 'フェリー代']
public_expenses = [jr, taxi, ferry]

for i in range(len(public_expenses)):
    if public_expenses[i]:
        public_expenses_len += 1



print(f'jr料金 : {jr}  円')
print(f'タクシー料金 : {taxi}  円')
print(f'フェリー代 : {ferry}  円')


# 単位判定変換
for i in range(dts_list_length):
    if dts_list[i][3] == '時間':
        dts_float_conver = float(dts_list[i][2])
        dts_int_conver = int(dts_float_conver)
        dts_list[i][2] = str(dts_int_conver) + 'H'
    if dts_list[i][3] == '件':
        dts_list[i][2] = str(dts_int_conver) + '日'



# job-Nomber
job_No = tables[0][1][1]


# total-amount
total_amount = tables[0][1][2]


total_amount = convert_int(total_amount)


tax_amount = total_amount * tax
tax_total_amount = tax_amount + total_amount
tax_total_amount = math.ceil(tax_total_amount)

print('-----------------------------------------------------------------------')
print(f'JOB-No : {job_No}')

print(f'小計合算金額 : {total_amount}  円')

print(f'税込み合計金額 : {tax_total_amount}  円')
print('-----------------------------------------------------------------------')

print()

print('-----------------------------------------------------------------------')

print('書き込み開始')

print('-----------------------------------------------------------------------')

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
    par = '{:.1%}'.format((y / dts_list_length))
    print()
    
    print(f'  {y+1}  /  {dts_list_length}  {par}')

    print()
    
    cell_len += 2
    if dts_list[y][1].startswith('宿泊'):
        sb[f'B{cell_len}'] = dts_list[y][1]
        sb[f'G{cell_len}'] = str_cell_data[0]
        sb[f'H{cell_len}'] = str_cell_data[1]
        dts_amount_int_1 = convert_int(dts_list[y][5])
        sb[f'I{cell_len}'] = dts_amount_int_1
        sb[f'J{cell_len}'] = '内税'


    elif dts_list[y][1].startswith('旅費'):
        for i in range(3):
            if public_expenses[i]:
                sb[f'B{cell_len}'] = public_expense_vehicles[i]
                sb[f'G{cell_len}'] = str_cell_data[0]
                sb[f'H{cell_len}'] = str_cell_data[1]
                sb[f'I{cell_len}'] = int(public_expenses[i])
                sb[f'J{cell_len}'] = '内税'
                cell_len += 2


    else:
        sb[f'B{cell_len}'] = dts_list[y][1]
        sb[f'G{cell_len}'] = dts_list[y][2]
        dts_unit_price_int = convert_int(dts_list[y][4])
        sb[f'H{cell_len}'] = dts_unit_price_int
        dts_amount_int_0 = convert_int(dts_list[y][5])
        sb[f'I{cell_len}'] = dts_amount_int_0



wb.save(excel_path)
print('-----------------------------------------------------------------------')

print('書き込み終了')

print('-----------------------------------------------------------------------')

print('-----------------------------------------------------------------------')

print('請求書作成処理　終了')

print('-----------------------------------------------------------------------')

