import pandas as pd
import json

file_path = r'd:\test ch\閱讀教學策略查詢系統\學習扶助閱讀測驗試題分析.xlsx'
xl = pd.ExcelFile(file_path)

# Load sheets
df_grades = {
    '三年級': pd.read_excel(file_path, sheet_name='三年級'),
    '四年級': pd.read_excel(file_path, sheet_name='四年級'),
    '五年級': pd.read_excel(file_path, sheet_name='五年級'),
    '六年級': pd.read_excel(file_path, sheet_name='六年級')
}
df_trans = pd.read_excel(file_path, sheet_name='學習遷移題目')

# Combine original questions into one dataframe for easy lookup
df_orig = pd.concat(df_grades.values(), ignore_index=True)

# Get the first 3 transfer questions for testing
test_cases = []
for i in range(3):
    trans_row = df_trans.iloc[i]
    orig_idx = trans_row['對應題目ID']
    # Find matching row
    matches = df_orig[df_orig['題目ID'] == orig_idx]
    if len(matches) == 0:
        continue
    orig_row = matches.iloc[0]
    
    test_cases.append({
        'original': {
            'id': str(orig_idx),
            'text': str(orig_row['文本內容'])[:150] + '...',
            'question': str(orig_row['題目']),
            'options': f"A: {orig_row['選項A']}, B: {orig_row['選項B']}, C: {orig_row['選項C']}, D: {orig_row['選項D']}",
            'answer': str(orig_row['正確答案']),
            'cognitive': str(orig_row['認知歷程']),
            'strategy': str(orig_row['教學策略名稱'])
        },
        'transfer': {
            'text': str(trans_row['遷移文本內容'])[:150] + '...',
            'question': str(trans_row['遷移題目']),
            'options': f"A: {trans_row['遷移選項A']}, B: {trans_row['遷移選項B']}, C: {trans_row['遷移選項C']}, D: {trans_row['遷移選項D']}",
            'answer': str(trans_row['遷移正確答案']),
            'explanation': str(trans_row['遷移解析'])
        }
    })

print(json.dumps(test_cases, ensure_ascii=False, indent=2))
