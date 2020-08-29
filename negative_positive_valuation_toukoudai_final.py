# モジュールのインポート
import re
import csv
import time
import pandas as pd
import matplotlib.pyplot as plt
import MeCab
import random
import openpyxl
import numpy

pn_df = pd.read_excel(r'C:\Users\VARNG\Desktop\scrapping\pn_table.xlsx',
                    encoding='utf-8',
                    names=('Word','Reading','POS', 'PN')
                   )

# PN Tableをデータフレームからdict型に変換しておく
word_list = list(pn_df['Word'])
pn_list = list(pn_df['PN'])  # 中身の型はnumpy.float64
pn_dict = dict(zip(word_list, pn_list))

# MeCabインスタンス作成
m = MeCab.Tagger('')  # 指定しなければIPA辞書

def get_diclist(text):
    parsed = m.parse(text)      # 形態素解析結果（改行を含む文字列として得られる）
    lines = parsed.split('\n')  # 解析結果を1行（1語）ごとに分けてリストにする
    lines = lines[0:-2]         # 後ろ2行は不要なので削除
    diclist = []
    for word in lines:
        l = re.split('\t|,',word)  # 各行はタブとカンマで区切られてるので
        d = {'Surface':l[0], 'POS1':l[1], 'POS2':l[2], 'BaseForm':l[7]}
        diclist.append(d)
    return(diclist)


# 形態素解析結果の単語ごとdictデータにPN値を追加する関数
def add_pnvalue(diclist_old):
    diclist_new = []
    for word in diclist_old:
        base = word['BaseForm']        # 個々の辞書から基本形を取得
        if base in pn_dict:
            pn = float(pn_dict[base])  # 中身の型があれなので
        else:
            pn = 'notfound'            # その語がPN Tableになかった場合
        word['PN'] = pn
        diclist_new.append(word)
    return(diclist_new)

def get_pnmean(diclist):
    pn_list = []
    for word in diclist:
        pn = word['PN']
        if pn != 'notfound':
            pn_list.append(pn)  # notfoundだった場合は追加もしない
    if len(pn_list) > 0:        # 「全部notfound」じゃなければ
        pnmean = numpy.mean(pn_list)
    else:
        pnmean = 0              # 全部notfoundならゼロにする
    return(pnmean)

tw_df = pd.read_excel(r'C:\Users\VARNG\Desktop\scrapping\SNS\jptestfor0.xlsx', encoding='utf-8')

print(tw_df)

# 一応、本文テキストから改行を除いておく（最初にやれ）
pnmeans_list = []
for tw in tw_df['comment']:
    dl_old = get_diclist(tw)
    dl_new = add_pnvalue(dl_old)
    pnmean = get_pnmean(dl_new)
    pnmeans_list.append(pnmean)

text_list = list(tw_df['comment'])
for i in range(len(text_list)):
    text_list[i] = text_list[i].replace('\n', ' ')


# ツイートID、本文、PN値を格納したデータフレームを作成
aura_df = pd.DataFrame({
                        'text':text_list,
                        'PN':pnmeans_list,
                       },
                       columns=['text', 'PN']
                      )


# PN値の昇順でソート
aura_df = aura_df.sort_values(by='PN', ascending=True)


# CSVを出力（ExcelでみたいならUTF8ではなくShift-JISを指定すべき）
aura_df.to_excel('C:/Users/VARNG/Desktop/scrapping/20200829_finance_hack/pn_result.xlsx',\
                index=None,\
                encoding='shift-jis',\
               )
