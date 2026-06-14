import os
import re
import sys
import tempfile
import subprocess
from datetime import datetime

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
except Exception as e:
    print('Missing python-pptx. Please install dependencies: pip install -r requirements.txt')
    raise

MERMAID_REGEX = re.compile(r"```mermaid\n(.*?)```", re.S)


def detect_theme(text):
    text_lower = text.lower()
    tech_keywords = ['engineer', 'technical', 'api', 'architecture', 'deploy', 'devops', 'infra']
    corp_keywords = ['business', 'strategy', 'market', 'sales', 'finance']
    if any(k in text_lower for k in tech_keywords):
        return 'technical'
    if any(k in text_lower for k in corp_keywords):
        return 'corporate'
    return 'technical'


def parse_markdown(text):
    title = None
    lines = text.strip().splitlines()
    for line in lines:
        if line.strip():
            # first non-empty line as title if starts with # or not
            if line.lstrip().startswith('#'):
                title = line.lstrip('#').strip()
            else:
                title = line.strip()
            break
    # split into sections by '## '
    parts = re.split(r'^##\s+', text, flags=re.M)
    sections = []
    if parts:
        if parts[0].strip():
            # leading content without section header
            sections.append(('Overview', parts[0].strip()))
        for p in parts[1:]:
            header, *body = p.splitlines()
            sections.append((header.strip(), '\n'.join(body).strip()))
    return title or 'Presentation', sections


def render_mermaid(code, out_png):
    # Try to render using mmdc (Mermaid CLI). If unavailable, raise.
    with tempfile.NamedTemporaryFile('w', suffix='.mmd', delete=False) as f:
        f.write(code)
        mmd_in = f.name
    # Try common commands
    cmds = [ ['mmdc', '-i', mmd_in, '-o', out_png], ['npx', 'mmdc', '-i', mmd_in, '-o', out_png] ]
    for cmd in cmds:
        try:
            subprocess.check_call(cmd)
            os.unlink(mmd_in)
            return True
        except Exception:
            continue
    os.unlink(mmd_in)
    return False


def create_pptx(title, sections, mermaid_blocks, theme):
    prs = Presentation()
    # set basic fonts/colors based on theme
    if theme == 'technical':
        bg_color = RGBColor(18, 18, 18)
        title_color = RGBColor(230, 230, 230)
        bullet_color = RGBColor(200, 200, 200)
    else:
        bg_color = RGBColor(255, 255, 255)
        title_color = RGBColor(20, 40, 80)
        bullet_color = RGBColor(50, 50, 50)

    # Title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    subtitle = slide.placeholders[1]
    subtitle.text = f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}'

    # Agenda if multiple sections
    if len(sections) > 1:
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = 'Agenda'
        body = slide.shapes.placeholders[1].text_frame
        for sec, _ in sections:
            p = body.add_paragraph()
            p.text = sec
            p.level = 0

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    os.makedirs(assets_dir, exist_ok=True)

    # place content slides
    mermaid_idx = 0
    for sec_title, sec_body in sections:
        # check for mermaid blocks in body
        mermaids = MERMAID_REGEX.findall(sec_body)
        if mermaids:
            for md in mermaids:
                # attempt to render
                img_name = f'mermaid_{mermaid_idx}.png'
                out_png = os.path.join(assets_dir, img_name)
                ok = render_mermaid(md, out_png)
                if ok and os.path.exists(out_png):
                    layout = prs.slide_layouts[5]  # blank
                    slide = prs.slides.add_slide(layout)
                    left = Inches(1)
                    top = Inches(1.2)
                    width = Inches(8)
                    slide.shapes.add_picture(out_png, left, top, width=width)
                    # caption
                    tx = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.7))
                    tf = tx.text_frame
                    tf.text = sec_title
                else:
                    # fallback slide
                    layout = prs.slide_layouts[1]
                    slide = prs.slides.add_slide(layout)
                    slide.shapes.title.text = f'Diagram: needs review — {sec_title}'
                    body = slide.shapes.placeholders[1].text_frame
                    body.text = 'Mermaid diagram could not be rendered; source included in notes.'
                    slide.notes_slide.notes_text_frame.text = md
                mermaid_idx += 1
        else:
            # normal text -> bullets
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = sec_title
            tf = slide.shapes.placeholders[1].text_frame
            # split into paragraphs
            paras = re.split(r'\n\n+', sec_body.strip())
            for ptxt in paras:
                if not ptxt.strip():
                    continue
                p = tf.add_paragraph()
                # use first line as bold maybe
                lines = ptxt.strip().splitlines()
                p.text = lines[0].strip()
                p.level = 0
                # additional lines become sub-bullets
                for sub in lines[1:]:
                    sp = tf.add_paragraph()
                    sp.text = sub.strip()
                    sp.level = 1

    out_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(out_dir, exist_ok=True)
    fname = f\"{re.sub(r'[^0-9a-zA-Z]+','_', title)[:50]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx\"
    out_path = os.path.join(out_dir, fname)
    prs.save(out_path)
    return out_path


def main(path):
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    title, sections = parse_markdown(text)
    theme = detect_theme(text)
    out = create_pptx(title, sections, None, theme)
    print('Saved PPTX to', out)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python agent_runner.py <input.md>')
        sys.exit(1)
    main(sys.argv[1])