import re
import sys

with open('README.md', 'rb') as f:
    data = f.read()
if data.startswith(b'\xef\xbb\xbf'):
    data = data[3:]

# Build mapping: corrupted utf-8 sequence -> original byte
mapping = {}

# Latin-1 range (0xA0-0xFF)
for b in range(0xA0, 0x100):
    mapping[chr(b).encode('utf-8')] = bytes([b])

# C1 control range (0x80-0x9F) for bytes NOT in Windows-1252
c1_exclude = {0x80, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
              0x8a, 0x8b, 0x8c, 0x8e, 0x91, 0x92, 0x93, 0x94, 0x95,
              0x96, 0x97, 0x98, 0x99, 0x9a, 0x9b, 0x9c, 0x9e, 0x9f}
for b in range(0x80, 0xA0):
    if b not in c1_exclude:
        mapping[chr(b).encode('utf-8')] = bytes([b])

# Windows-1252 specials
win1252 = {
    0x80: '\u20ac', 0x82: '\u201a', 0x83: '\u0192',
    0x84: '\u201e', 0x85: '\u2026', 0x86: '\u2020',
    0x87: '\u2021', 0x88: '\u02c6', 0x89: '\u2030',
    0x8a: '\u0160', 0x8b: '\u2039', 0x8c: '\u0152',
    0x8e: '\u017d',
    0x91: '\u2018', 0x92: '\u2019', 0x93: '\u201c',
    0x94: '\u201d', 0x95: '\u2022', 0x96: '\u2013',
    0x97: '\u2014', 0x98: '\u02dc', 0x99: '\u2122',
    0x9a: '\u0161', 0x9b: '\u203a', 0x9c: '\u0153',
    0x9e: '\u017e', 0x9f: '\u0178',
}
for b, ch in win1252.items():
    mapping[ch.encode('utf-8')] = bytes([b])

print(f'Mapping entries: {len(mapping)}')

# Process file byte by byte
result = bytearray()
i = 0
fix_count = 0
while i < len(data):
    matched = False
    for sl in [3, 2]:
        if i + sl <= len(data):
            chunk = bytes(data[i:i+sl])
            if chunk in mapping:
                result.extend(mapping[chunk])
                i += sl
                matched = True
                fix_count += 1
                break
    if not matched:
        result.append(data[i])
        i += 1

print(f'Fixes applied: {fix_count}')

# Decode
fixed = result.decode('utf-8')
print(f'Size: {len(data)} -> {len(fixed.encode("utf-8"))} bytes')

# Fix duplicate separator
sep = 'https://capsule-render.vercel.app/api?type=rect&color=gradient&customColorList=2&height=1&width=100%&section=header&animation=fadeIn'
double_sep = f"<img src='{sep}' width='100%' />\n\n<br>\n<img src='{sep}' width='100%' />"
count = fixed.count(double_sep)
if count > 0:
    fixed = fixed.replace(double_sep, f"<img src='{sep}' width='100%' />\n<br>", count)
    print(f'Removed {count} duplicate separator(s)')
else:
    print('No duplicate separator found')

# Fix data URIs
icon_order = [
    'apex', 'wordpress', 'lwc', 'javafx', 'winforms', 'swing', 'mvc',
    'jwt', 'oauth', 'bcrypt', 'jdbc', 'restapi', 'mern', 'websockets', 'langchain', 'passport',
    'railway', 'helm',
    'pandas', 'numpy', 'scikitlearn', 'mlops', 'recharts', 'promptengineering', 'agenticai',
    'openai', 'langchain', 'huggingface', 'anthropic', 'llms', 'agile',
    'salesforce'
]
uri_pattern = re.compile(r"<img src='data:image/svg\+xml;base64,[^']+' width='48' height='48' />")
matches = uri_pattern.findall(fixed)
if len(matches) == len(icon_order):
    c = [0]
    def repl(m):
        i = c[0]; c[0] += 1
        return f"<img src='icons/{icon_order[i]}.svg' width='48' height='48' />"
    fixed = uri_pattern.sub(repl, fixed)
    print(f'Replaced {len(matches)} data URIs')
else:
    print(f'Data URI match count: {len(matches)} (expected {len(icon_order)})')
    # Check if already replaced
    existing = len(re.findall(r"src='icons/", fixed))
    print(f'  Already have {existing} icon references')

# Write
with open('README.md', 'wb') as f:
    f.write(fixed.encode('utf-8'))

# Verify
with open('README.md', 'rb') as f:
    v = f.read()

correct = 0
for i in range(len(v)-3):
    if v[i]==0xF0 and v[i+1]==0x9F:
        correct += 1

repl_chars = 0
for i in range(len(v)-2):
    if v[i]==0xEF and v[i+1]==0xBF and v[i+2]==0xBD:
        repl_chars += 1

corrupted = 0
for i in range(len(v)-3):
    if v[i]==0xC3 and v[i+1]==0xB0 and v[i+2]==0xC5:
        corrupted += 1

print(f'Correct emoji starts (F0 9F): {correct}')
print(f'Replacement chars (EF BF BD): {repl_chars}')
print(f'Remaining corrupted (C3 B0 C5): {corrupted}')
print('DONE - file is clean!' if repl_chars == 0 and corrupted == 0 else 'ISSUES REMAIN')
