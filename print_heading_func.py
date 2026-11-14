with open('generate_full_paper.py', 'r', encoding='utf-8') as f:
    text = f.read()

start = text.find('def add_chapter_heading')
print(text[start:start+500])

