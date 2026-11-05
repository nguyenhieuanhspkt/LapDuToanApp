import re

with open('pdf_extracted_text.txt', 'r', encoding='utf-8') as f:
    text = f.read()

# Let's search for E6C, E6B, E6BOP
our_codes = [
    'E6C.3.05', 'E6C.3.01', 'E6C.3.02', 'E6C.3.04', 'E6C.3.03',
    'E6C.3.06', 'E6C.1.31', 'E6B.1.13.01', 'E6C.2.04', 'E6C.1.23',
    'E6C.1.18', 'E6BOP.1.03', 'E6BOP.1.01', 'E6BOP.2.01', 'E6BOP.2.16',
    'E6BOP.5.03', 'E6BOP.6.01'
]

print("=== SEARCHING OUR CODES IN PL QD 177 ===")
for c in our_codes:
    # search pattern
    pos = 0
    matches = []
    while True:
        idx = text.find(c, pos)
        if idx == -1:
            break
        # grab surrounding text
        snippet = text[max(0, idx - 100):min(len(text), idx + 250)]
        matches.append(snippet.strip())
        pos = idx + len(c)
    
    print(f"\nCode: {c} -> Found {len(matches)} occurrences")
    for m in matches[:3]:
        # print one-line summary
        cleaned = ' '.join(m.split())
        print(f"   Snippet: {cleaned}")
