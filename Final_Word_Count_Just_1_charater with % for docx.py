import sys
import importlib.util

# 檢查模組是否已安裝
def is_module_installed(module_name):
    spec = importlib.util.find_spec(module_name)
    return spec is not None

if is_module_installed("docx"):
    from docx import Document
else:
    print("警告：本環境中尚未安裝 python-docx 套件，請先安裝該套件。")
    sys.exit(1)

from collections import Counter
import pandas as pd
import os

# 定義要刪除的字元
drop_dict = [
    u'\u00A0',  # 新增不換行空白字元 U+00A0
    u'\uFF0E',  # 新增不換行空白字元 U+FF0E
    u'，', u'\n', u'。', u'、', u'：', u'(', u')', u'[', u']', u'.', u',', u' ',
    u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9', u'0',
    u'-', u'—', u'"', u'*', u'/', u':', u'@', u'～', u'《', u'》',
    u'『', u'』', u'+', u'○', u'◎', u'０', u'〇', u'２', u'4', u'５',
    u'６', u'８', u'9', u'a', u'A', u'B', u'b', u'c', u'C', u'd',
    u'e', u'E', u'f', u'F', u'g', u'h', u'H', u'i', u'j', u'k',
    u'l', u'L', u'm', u'M', u'n', u'N', u'o', u'p', u'P', u'Q',
    u'q', u'r', u'R', u's', u'S', u't', u'T', u'u', u'U', u'v',
    u'V', u'w', u'x', u'y', u'Y', u'Z', u'G', u'D', u'I', u'J',
    u'K', u'O', u'X', u';', u'&', u'%', u'$', u'、', u'。', u'，',
    u'（', u'）', u'；', u'４', u'\u3000', u'”', u'“', u'？', u'?',
    u'！', u'‘', u'’', u'…', u'「', u'」', u'」', u'[', u']', u'．', u'【', u'】'
]

# 計算字元出現次數的函式，讀取 docx 檔案
def count_chars(file_path):
    # 讀取 docx 檔案
    doc = Document(file_path)
    # 將所有段落的文字合併成一個字串
    text = ''.join([para.text for para in doc.paragraphs])
    # 移除不需要的字元
    for char in drop_dict:
        text = text.replace(char, '')
    char_counts = Counter(text)
    return char_counts

# 取得使用者輸入的檔案路徑（請輸入 docx 檔案的路徑）
file_path = input("請輸入您的 docx 檔案路徑: ")

# 呼叫函式取得字元計數
char_counts = count_chars(file_path)

# 將字典轉換為 DataFrame，初始欄位分別為 'Character' 與 'Count'
df = pd.DataFrame(list(char_counts.items()), columns=['Character', 'Count'])

# 計算所有字元的總和
total_count = df['Count'].sum()

# 新增一個欄位，顯示該字元在全部字數中所佔的百分比（保留2位小數），並將欄位標題改為 "%"
df["%"] = (df['Count'] / total_count * 100).round(2)

# 定義優先排序的字元清單（順序依下）
priority_chars = ['之', '也', '故', '此', '如', '所', '等', '說', '中', '而', '是', '於']

# 若優先字元未出現在計數結果中，則補上並設定計數為 0
for ch in priority_chars:
    if ch not in df['Character'].values:
        df = df.append({'Character': ch, 'Count': 0, '%': 0.00}, ignore_index=True)

# 新增輔助欄位以進行排序：
# 若字元在 priority_chars 中，則設定 sort_group 為 0，並以 priority_chars 中的位置排序
# 其他字元則設定 sort_group 為 1，並以 Count 進行降冪排序（以負值處理）
df['sort_group'] = df['Character'].apply(lambda x: 0 if x in priority_chars else 1)
df['sort_order'] = df.apply(lambda row: priority_chars.index(row['Character']) if row['Character'] in priority_chars else -row['Count'], axis=1)

# 根據輔助欄位進行排序
df_sorted = df.sort_values(by=['sort_group', 'sort_order'])

# 移除輔助欄位
df_sorted = df_sorted.drop(columns=['sort_group', 'sort_order'])

# 新增 TOTAL 列（放在最後一列）
df_total = pd.DataFrame({
    'Character': ['TOTAL'],
    'Count': [total_count],
    '%': [100.00]
})

# 合併排序後的結果與 TOTAL 列
df_final = pd.concat([df_sorted, df_total], ignore_index=True)

# 將「Count」轉為整數型態，"%" 轉為浮點型態
df_final['Count'] = pd.to_numeric(df_final['Count'], downcast='integer')
df_final["%"] = pd.to_numeric(df_final["%"], downcast='float')

# 根據輸入檔案的檔名產生輸出檔案名稱（同一目錄下，副檔名改為 .xlsx）
input_basename = os.path.basename(file_path)         # 例如 "sample.docx"
input_name, _ = os.path.splitext(input_basename)       # 取得 "sample"
output_file = os.path.join(os.path.dirname(file_path), input_name + ".xlsx")

# 將結果輸出到指定路徑
df_final.to_excel(output_file, index=False)

print(f"結果已儲存到 '{output_file}'")
