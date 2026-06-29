import docx

doc = docx.Document(r'C:\Users\786 COMPUTERS\Downloads\resume\Resume_Syed_Abdullah_Yaqoob_v3.docx')
for i, p in enumerate(doc.paragraphs):
    pf = p.paragraph_format
    print(f'=== PARA {i} ===')
    print(f'spacing_before={pf.space_before}')
    print(f'spacing_after={pf.space_after}')
    print(f'line_spacing={pf.line_spacing}')
    print(f'line_spacing_rule={pf.line_spacing_rule}')
    print(f'text: """{p.text}"""')
    print()
