import os
import xlrd
import pprint
import requests

#######################################
# リストファイル関連 設定
#######################################
fname = 'search_list.xls'                       # リストファイル名
wb = xlrd.open_workbook( fname )                # xlsファイルのBookオブジェクトを取得
sheet = wb.sheet_by_name( 'stocks' )            # 指定シートを取得

#######################################
# 出力ディレクトリ 作成
#######################################
cnt = 1                                         # 作成ディレクトリ名を変更するための数字
new_dir_path = 'jpg_files'                      # 作成するディレクトリ名
while True:
    try:
        os.mkdir(new_dir_path)                  # ディレクトリを作成する
        break                                   # ディレクトリ作成が成功すれば抜ける
    except FileExistsError:                     # ディレクトリが存在していた場合
        new_dir_path = 'jpg_files' + str(cnt)   # ディレクトリ名を変更(末尾に番号付加)
        cnt += 1                                # 末尾に付加用の数字+1

#######################################
# スクレイピング 設定
#######################################
# 時間　1日: 1d、1週: 5d、1ヶ月: 1m、3ヶ月: 3m、6ヶ月: 6m、1年: 1y、2年: 2y、5年: 5y、10年: ay
paramTime = 'ay'                                # ※※要設定※※
# 画像サイズ　標準: m、大: n
paramSize = 'n'                                 # ※※要設定※※
# 基本URL
urlShort1 = 'https://chart.yahoo.co.jp/?code='  # この後ろに「銘柄番号」
urlShort2 = '.T&tm='                            # この後ろに「時間」
urlShort3 = '&vip=off'
urlLong1 = 'https://chart.yahoo.co.jp/?code='   # この後ろに「銘柄番号」
urlLong2 = '.T&tm='                             # この後ろに「時間」
urlLong3 = '&type=c&log=off&size='              # この後ろに「画像サイズ」
urlLong4 = '&over=m65,m130,s&add=v&comp='

# 全画像ファイル出力
lineno = 1
while lineno < sheet.nrows:                     # 最終行まで
    cell = sheet.cell( lineno, 0 )              # セルを読む
    if cell.ctype == 0:                         # セルが空白の場合
        break                                   # ループを抜ける
    if cell.value > 1000 or cell.value < 10000: # セルが正常範囲

        # 出力ファイル名を作成
        if os.name == 'nt':     # OSがWindowsの場合
            output_path = os.getcwd() + '\\' + new_dir_path + '\\' + str(int(cell.value)) + '.png'
        else:                   # OSがWindows以外の場合(Linuxを想定)
            output_path = os.getcwd() + '/' + new_dir_path + '/' + str(int(cell.value)) + '.png'

        # 出力ファイルを開く
        with open( output_path, 'wb') as f:
            # URL 作成
            if paramTime == '1d'  or paramTime == '5d':     # URL が Short かの判定
                url = urlShort1 + str(int(cell.value)) + urlShort2 + paramTime + urlShort3
            else:
                url = urlLong1 + str(int(cell.value)) + urlLong2 + paramTime + urlLong3 + paramSize + urlLong4

            re = requests.get( url )
            f.write( re.content )               # ファイル出力
    lineno += 1                                 # 行+1


