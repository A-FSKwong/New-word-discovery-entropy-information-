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

# 詢問使用者最小次數、最小支持度、最大熵篩選
min_count = int(input("Enter the minimum count: "))
min_support = int(input("Enter the minimum support: "))
min_s = int(input("Enter the minimum entropy: "))

# 讀取文章內容
with open(file_path, 'r', encoding='utf-8') as f:
    s = f.read()

# 定義要去掉的標點符號
drop_dict = [u'，', u'\n', u'。', u'、', u'：', u'(', u')', u'[', u']', u'.', u',', u' ', 
             u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9', u'0',
             u'-', u'—', u'"', u'*', u'/', u':', u'@', u'～', u'《', u'》',
             u'『', u'』', u'+', u'○', u'◎', u'０', u'〇', u'２', u'4', u'５',
             u'６', u'８', u'9', u'a', u'A', u'B', u'b', u'c', u'C', u'd',
             u'e', u'E', u'f', u'F', u'g', u'h', u'H', u'i', u'j', u'k',
             u'l', u'L', u'm', u'M', u'n', u'N', u'o', u'p', u'P', u'Q',
             u'q', u'r', u'R', u's', u'S', u't', u'T', u'u', u'U', u'v',
             u'V', u'w', u'x', u'y', u'Y', u'Z', u'G', u'D', u'I', u'J',
             u'K', u'O', u'X', u';', u'&', u'%', u'$', u'、', u'。', u'，',u'（', u'）', u'；', u'４', 
             u'\u3000', u'”', u'“', u'？', u'?', u'！', u'‘', u'’', u'…', u'「', u'」']


for i in drop_dict: # 去掉標點符號
    s = s.replace(i, '')

# 自定義正則表達式的詞典
myre = {i: f'({"." * i})' for i in range(min_len, max_len + 1)}

# 初始化結果列表
t = [pd.Series(list(s)).value_counts()]  # 逐字統計
tsum = t[0].sum()  # 統計總字數
rt = []  # 保存篩選結果

# 生成詞語並篩選
for m in range(min_len, max_len + 1):
    print(f'正在生成{m}字詞...')
    if m > 1:
        t.append(pd.Series(dtype=object))  # 初始化為空的Series
    for i in range(m):  # 生成所有可能的m字詞
        found_words = pd.Series(re.findall(myre[m], s[i:]))
        if m == 1:
            # 如果是一字詞，則不進行後續處理
            continue
        else:
            # Concatenate found words to the existing Series
            t[m - 1] = pd.concat([t[m - 1], found_words], ignore_index=True)

    if m > 1:
        t[m - 1] = t[m - 1].value_counts()  # 逐詞統計
        t[m - 1] = t[m - 1][t[m - 1] > min_count]  # 最小次數篩選
        tt = t[m - 1][:]
        for k in range(m - 1):
            qq = np.array(list(map(lambda ms: tsum * tt[ms] / t[m - 2 - k][ms[:m - 1 - k]] / t[k][ms[m - 1 - k:]], tt.index))) > min_support  # 最小支持度篩選
            tt = tt[qq]
        rt.append(tt.index)

# 信息熵計算函數
def cal_S(sl):
    return -((sl / sl.sum()).apply(log) * sl / sl.sum()).sum()

# 進行最大熵篩選
for i in range(2, max_len + 1):
    print(f'正在進行{i}字詞的最大熵篩選({len(rt[i - 2])})...')
    pp = []  # 保存左右鄰居結果
    for j in range(i + 2):
        
        pp = pp + re.findall(f'(.){myre[i]}(.)', s[j:])

    pp = pd.DataFrame(pp).set_index(1).sort_index()  # 排序
    index = np.sort(np.intersect1d(rt[i - 2], pp.index))  # 取交集

    # 左右熵篩選
    index = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[0][s]).value_counts()), index))) > min_s]
    rt[i - 2] = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[2][s]).value_counts()), index))) > min_s]

# 輸出處理
for i in range(len(rt)):
    t[i + 1] = t[i + 1][rt[i]]
    t[i + 1] = t[i + 1].sort_values(ascending=False)

# 獲取原始文件的目錄
dir_path = os.path.dirname(file_path)

# 將結果保存到Excel檔案
output_filename = f'result_entropy{min_s}_min_count{min_count}_min_support{min_support}.xlsx'
output_path = os.path.join(dir_path, output_filename)
with pd.ExcelWriter(output_path) as writer:
    for i, df in enumerate(t, start=1):  # Start at sheet '1 character words'
        df_sorted = df.sort_values(ascending=False)
        df_to_save = pd.DataFrame(df_sorted)
        df_to_save.columns = ['Words count']
        df_to_save.index.name = 'Words'
        df_to_save.to_excel(writer, sheet_name=f'{i} character words')

print(f"結果已保存至'{output_filename}'，與源文件在同一目錄下。")
