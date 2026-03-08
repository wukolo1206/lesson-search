import json
with open('bad_options.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for item in data[5:25]:
    article = item['article'].replace('\n', ' ').replace('\r', '')
    title = article[:30]
    q = item['q'].replace('\n', ' ').replace('\r', '')
    print(f"Row {item['row']} | {title} | {q}")
