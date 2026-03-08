import pandas as pd
import json

file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
df_orig = pd.read_excel(file_path, sheet_name='三年級')
df_trans = pd.read_excel(file_path, sheet_name='學習遷移題目')

# Let's look at the absolute first row of trans sheet
trans_row = df_trans.iloc[0]
orig_title = trans_row['原始案例標題']

matches = df_orig[df_orig['文本標題'] == orig_title]
orig_row = matches.iloc[0]

test_cases = {
    'original': {
        'title': str(orig_row['文本標題']),
        'question': str(orig_row['完整題目']),
        'strategy': str(orig_row['教學策略名稱'])
    },
    'transfer': {
        'text_snippet': str(trans_row['遷移文本內容'])[:20] + '...',
        'question': str(trans_row['遷移題目'])
    }
}
print(json.dumps(test_cases, ensure_ascii=False, indent=2))
