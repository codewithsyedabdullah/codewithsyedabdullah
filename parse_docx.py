import docx
from docx.shared import Pt, Emu

doc = docx.Document(r"C:\Users\786 COMPUTERS\Downloads\resume\Resume_Syed_Abdullah_Yaqoob_v3.docx")

for i, para in enumerate(doc.paragraphs):
    print(f"{'='*80}")
    print(f"PARAGRAPH INDEX: {i}")
    print(f"{'='*80}")
    print(f"Full text (repr): {repr(para.text)}")
    print(f"Text length: {len(para.text)}")
    print(f"Style: {para.style.name if para.style else 'None'}")

    pf = para.paragraph_format
    ls = pf.line_spacing
    ls_rule = pf.line_spacing_rule
    print(f"line_spacing: {ls}")
    print(f"line_spacing_rule: {ls_rule}")
    print(f"space_before: {pf.space_before}")
    print(f"space_after: {pf.space_after}")

    text_stripped = para.text.strip()
    if text_stripped == '' or text_stripped == '\t' or text_stripped == ' ':
        classification = 'SPACER (empty/whitespace only)'
    elif text_stripped == '_x000D_':
        classification = 'SEPARATOR (carriage return)'
    elif text_stripped in ('-', '_', '___', '---', '__________', '______________________________________________'):
        classification = 'SEPARATOR (dash/underline)'
    else:
        classification = 'NORMAL'
    print(f"Classification: {classification}")

    runs = para.runs
    print(f"Runs count: {len(runs)}")
    for j, run in enumerate(runs):
        print(f"  Run[{j}]:")
        print(f"    text (repr): {repr(run.text)}")
        print(f"    bold: {run.bold}")
        print(f"    font_name: {run.font.name}")
        print(f"    font_size: {run.font.size}")
        print(f"    font_color: {run.font.color.rgb if run.font.color and run.font.color.rgb else 'None'}")
        print(f"    italic: {run.italic}")
        print(f"    underline: {run.underline}")
    print()
