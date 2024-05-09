import numpy as np  # 引入 numpy 庫，用於數據處理
import pandas as pd  # 引入 pandas 庫，用於數據分析
import re  # 引入 re 庫，用於正則表達式操作
from numpy import log  # 從 numpy 庫中引入 log 函數，用於計算信息熵
import os  # 引入 os 庫，用於操作系統路徑


# 詢問使用者來源文件的路徑
file_path = input("Enter the path of the source file: ")  # 獲取用戶輸入的文件路徑


# 詢問使用者詞語的最小和最大長度
min_len = int(input("Enter the minimum length of words: "))  # 獲取用戶輸入的詞語最小長度

max_len = int(input("Enter the maximum length of words: "))  # 獲取用戶輸入的詞語最大長度




# 詢問使用者 min_count，min_support 和 min_s

"""
min_count   錄取詞語最小出現次數
min_support 錄取詞語最低支持度，1代表著隨機組合
min_s       錄取詞語最低信息熵，越大說明越有可能獨立成詞

"""
min_count = int(input("Enter the minimum count: "))  # 獲取用戶輸入的詞語最小出現次數

min_support = int(input("Enter the minimum support: "))  # 獲取用戶輸入的詞語最低支持度

min_s = int(input("Enter the minimum entropy: "))  # 獲取用戶輸入的詞語最低信息熵




f = open(file_path, 'r', encoding='utf-8')  # 以讀取模式打開用戶指定的文件

s = f.read()  # 讀取文件內容為一個字符串


# 定義要去掉的標點字
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



for i in drop_dict:  # 對每一個要去掉的標點字
    s = s.replace(i, '')  # 在字符串中替換掉該標點字
	


# 為了方便調用，自定義了一個正則表達式的詞典
myre = {i:f'({"."*i})' for i in range(min_len, max_len+1)}


t=[]  # 初始化一個空列表，用於保存結果

t.append(pd.Series(list(s)).value_counts())  # 對字符串進行逐字統計

tsum = t[0].sum()  # 統計總字數

rt = []  # 初始化一個空列表，用於保存結果


for m in range(min_len, max_len+1):  # 對於每一個詞語長度
    print(u'正在生成%s字詞...'%m)  # 輸出提示信息

    t.append([])  # 在結果列表中添加一個空列表

    for i in range(m):  # 對於每一個位置
        t[m-1] = t[m-1] + re.findall(myre[m], s[i:])  # 生成所有可能的m字詞
    

    t[m-1] = pd.Series(t[m-1]).value_counts()  # 對詞語進行逐詞統計

    t[m-1] = t[m-1][t[m-1] > min_count]  # 對詞語進行最小次數篩選

    tt = t[m-1][:]  # 複製一份統計結果
	

    for k in range(m-1):  # 對於每一個位置
        qq = np.array(list(map(lambda ms: tsum*t[m-1][ms]/t[m-2-k][ms[:m-1-k]]/t[k][ms[m-1-k:]], tt.index))) > min_support  # 進行最小支持度篩選

        tt = tt[qq]  # 更新統計結果

    rt.append(tt.index)  # 將結果添加到結果列表中		
	
	
	


def cal_S(sl):  # 定義信息熵計算函數
    return -((sl/sl.sum()).apply(log)*sl/sl.sum()).sum()
	
	


for i in range(min_len, max_len+1):  # 對於每一個詞語長度
    print(u'正在進行%s字詞的最大熵篩選(%s)...'%(i, len(rt[i-2])))  # 輸出提示信息

    pp = []  # 初始化一個空列表，用於保存所有的左右鄰結果

    for j in range(i+2):  # 對於每一個位置
        pp = pp + re.findall('(.)%s(.)'%myre[i], s[j:])  # 生成所有可能的左右鄰結果

    pp = pd.DataFrame(pp).set_index(1).sort_index()  # 將結果轉換為數據框並排序

    index = np.sort(np.intersect1d(rt[i-2], pp.index))  # 進行交集操作
	


    # 下面兩句分別是左鄰和右鄰信息熵篩選
    index = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[0][s]).value_counts()), index))) > min_s]

    rt[i-2] = index[np.array(list(map(lambda s: cal_S(pd.Series(pp[2][s]).value_counts()), index))) > min_s]
	


# 下面都是輸出前處理
for i in range(len(rt)):
    t[i+1] = t[i+1][rt[i]]

    t[i+1] = t[i+1].sort_values(ascending = False)


# 獲取源文件的目錄
dir_path = os.path.dirname(file_path)


# 保存結果並輸出
with pd.ExcelWriter(os.path.join(dir_path, f'result_entropy{min_s}_min_count{min_count}_min_support{min_support}.xlsx')) as writer:  
    for i in range(len(t[1:])):
        df = pd.DataFrame(t[i+1])
        df.columns = ['Words count']
        df.index.name = 'Words'
        df.to_excel(writer, sheet_name=f'{i+2} character words')
		

print(f"Results saved to 'result_entropy{min_s}_min_count{min_count}_min_support{min_support}.xlsx' in the same directory as the source file.")
