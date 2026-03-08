import json
with open('bad_options.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

with open('batch4.txt', 'w', encoding='utf-8') as out:
    for item in data[55:95]:
        out.write(f"Row {item['row']}\n")
        out.write(f"Q: {item['q']}\n")
        out.write(f"Article (first 350 chars):\n")
        out.write(f"{item['article'][:350]}\n")
        out.write('='*60 + '\n')
