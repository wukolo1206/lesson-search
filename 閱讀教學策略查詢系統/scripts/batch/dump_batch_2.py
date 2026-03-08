import json
with open('bad_options.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

with open('batch2.txt', 'w', encoding='utf-8') as out:
    for item in data[5:25]:
        out.write(f"Row {item['row']}\n")
        out.write(f"Q: {item['q']}\n")
        out.write(f"Article (first 250 chars):\n")
        out.write(f"{item['article'][:250]}\n")
        out.write('='*60 + '\n')
