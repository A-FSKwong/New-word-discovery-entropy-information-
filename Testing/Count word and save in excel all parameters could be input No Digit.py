import numpy as np
import pandas as pd
import re
from numpy import log
import os

# 詢問使用者來源文件的路徑
file_path = input("Enter the path of the source file: ")

# 詢問使用者詞語的最小和最大長度
min_len = int(input("Enter the minimum length of words: "))
max_len = int(input("Enter the maximum length of words: "))


# 詢問使用者 min_count，min_support 和 min_s

"""
min_count   錄取詞語最小出現次數
min_support 錄取詞語最低支持度，1代表著隨機組合
min_s       錄取詞語最低信息熵，越大說明越有可能獨立成詞

"""
min_count = int(input("Enter the minimum count: "))
min_support = int(input("Enter the minimum support: "))
min_s = int(input("Enter the minimum entropy: "))



f = open(file_path, 'r', encoding='utf-8') #讀取文章
s = f.read() #讀取為一個字符串

#定義要去掉的標點字

drop_dict = [u'，', u'\n', u'。', u'、', u'：', u'(', u')', u'[', u']', u'.', u',', u' ', 
             u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9', u'0'
             u'\u3000', u'”', u'“', u'？', u'?', u'！', u'‘', u'’', u'…', u'「', u'」']

for i in drop_dict: #去掉標點字
    s = s.replace(i, '')

#為了方便調用，自定義了一個正則表達式的詞典
myre = {i:f'({"."*i})' for i in range(min_len, max_len+1)}

t=[] #保存結果用。

t.append(pd.Series(list(s)).value_counts()) #逐字統計
tsum = t[0].sum() #統計總字數
rt = [] #保存結果用

for m in range(min_len, max_len+1):
    print(u'正在生成%s字詞...'%m)
    t.append([])
    for i in range(m): #生成所有可能的m字詞
        t[m-1] = t[m-1] + re.findall(myre[m], s[i:])
    
    t[m-1] = pd.Series(t[m-1]).value_counts() #逐詞統計
    t[m-1] = t[m-1][t[m-1] > min_count] #最小次數篩選
    tt = t[m-1][:]
    for k in range(m-1):
        qq = np.array(list(map(lambda ms: tsum*t[m-1][ms]/t[m-2-k][ms[:m-1-k]]/t[k][ms[m-1-k:]], tt.index))) > min_support #最小支持度篩選。
        tt = tt[qq]
    rt.append(tt.index)

def cal_S(sl): #信息熵計算函數
    return -((sl/sl.sum()).apply(log)*sl/sl.sum()).sum()

for i in range(min_len, max_len+1):
    print(u'正在進行%s字詞的最大熵篩選(%s)...'%(i, len(rt[i-2])))
    pp = [] #保存所有的左右鄰結果
    for j in range(i+2):
        pp = pp + re.findall('(.)%s(.)'%myre[i], s[j:])
    pp = pd.DataFrame(pp).set_index(1).sort_index() #先排序，這個很重要，可以加快檢索速度
    index = np.sort(np.intersect1d(rt[i-2], pp.index)) #作交集

    #下面兩句分別是左鄰和右鄰信息熵篩選
    index = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[0][s]).value_counts()), index))) > min_s]
    rt[i-2] = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[2][s]).value_counts()), index))) > min_s]


#下面都是輸出前處理
for i in range(len(rt)):
    t[i+1] = t[i+1][rt[i]]
    t[i+1] = t[i+1].sort_values(ascending = False)

# Get the directory of the source file
dir_path = os.path.dirname(file_path)

#保存結果並輸出
with pd.ExcelWriter(os.path.join(dir_path, f'result_entropy{min_s}_min_count{min_count}_min_support{min_support}.xlsx')) as writer:  
    for i in range(len(t[1:])):
        df = pd.DataFrame(t[i+1])
        df.columns = ['Words count']
        df.index.name = 'Words'
        df.to_excel(writer, sheet_name=f'{i+2} character words')

print(f"Results saved to 'result_entropy{min_s}_min_count{min_count}_min_support{min_support}.xlsx' in the same directory as the source file.")