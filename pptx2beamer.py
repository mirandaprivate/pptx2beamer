#!/usr/bin/env python3
#
# pptx2beamer.py
# A script to extract content from PPTX and insert into a fixed Beamer template.
#

import sys
import os
import argparse
import zipfile
import shutil
import tempfile
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import re

def sanitize_for_latex(text):
    """Remove characters that are invalid for LaTeX command names."""
    sanitized = re.sub(r'[^a-zA-Z0-9]', '', text)
    if sanitized:
        return sanitized
    import hashlib
    return "layout" + hashlib.md5(text.encode('utf-8')).hexdigest()[:8]

def escape_latex(text):
    """Escape LaTeX special characters in text."""
    if text is None:
        return ""
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(ch, ch) for ch in text)

# --- XML Parsing Functions ---

def parse_slides_for_content(ppt_dir):
    """Parses actual slides to extract images and their positions and texts."""
    slides_data = []
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    slide_dir = ppt_dir / 'ppt' / 'slides'
    if not slide_dir.exists():
        return slides_data
        
    # Get sorted slide files
    slide_files = sorted(slide_dir.glob('slide*.xml'), 
                        key=lambda x: int(re.search(r'slide(\d+)\.xml', x.name).group(1)))
    
    for slide_file in slide_files:
        slide_num = re.search(r'slide(\d+)\.xml', slide_file.name).group(1)
        rels_file = slide_dir / '_rels' / f'{slide_file.name}.rels'
        
        slide_info = {
            'number': slide_num,
            'images': [],
            'title': f"Slide {slide_num}",
            'texts': []
        }
        
        if not rels_file.exists():
            slides_data.append(slide_info)
            continue
            
        # Parse relationships to resolve image IDs
        rel_tree = ET.parse(rels_file)
        rel_root = rel_tree.getroot()
        rels = {r.get('Id'): r.get('Target') for r in rel_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}
        
        try:
            tree = ET.parse(slide_file)
            root = tree.getroot()
            
            # Find title and body text
            for sp in root.findall('.//p:sp', ns):
                ph = sp.find('.//p:nvPr/p:ph', ns)
                ph_type = ph.get('type') if ph is not None else None

                texts = [t.text for t in sp.findall('.//a:t', ns) if t.text]
                if not texts:
                    continue
                text = "".join(texts).strip()
                if not text:
                    continue

                if ph_type in ('title', 'ctrTitle'):
                    if slide_info['title'].startswith("Slide "):
                        slide_info['title'] = text
                    continue
                if ph_type == 'subtitle':
                    # Subtitle is handled by title page parsing
                    continue

                slide_info['texts'].append(text)
            
            # Find all pictures on the slide
            for pic in root.findall('.//p:pic', ns):
                blip = pic.find('.//p:blipFill/a:blip', ns)
                if blip is None:
                    blip = pic.find('.//a:blip', ns)
                
                if blip is not None:
                    rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rId in rels:
                        img_path = rels[rId].split('/')[-1]
                        lower_img = img_path.lower()
                        if lower_img.endswith('.jfif'):
                            img_path = img_path[:-5] + ".jpg"
                        
                        # Extract coordinates (xfrm)
                        position = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                        xfrm = pic.find('.//a:xfrm', ns)
                        if xfrm is not None:
                            off = xfrm.find('.//a:off', ns)
                            ext = xfrm.find('.//a:ext', ns)
                            if off is not None:
                                position['x'] = int(off.get('x', 0))
                                position['y'] = int(off.get('y', 0))
                            if ext is not None:
                                position['width'] = int(ext.get('cx', 0))
                                position['height'] = int(ext.get('cy', 0))
                                
                        slide_info['images'].append({
                            'name': img_path,
                            'position': position
                        })
        except Exception:
            pass
            
        slides_data.append(slide_info)
        
    return slides_data

def parse_title_page_info(ppt_dir):
    """Extracts title/subtitle/author/institute/date from the first slide."""
    info = {
        'title': '',
        'subtitle': '',
        'author': '',
        'institute': '',
        'date': ''
    }
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }

    slide_path = ppt_dir / 'ppt' / 'slides' / 'slide1.xml'
    if not slide_path.exists():
        return info

    try:
        tree = ET.parse(slide_path)
        root = tree.getroot()

        for sp in root.findall('.//p:sp', ns):
            ph = sp.find('.//p:nvPr/p:ph', ns)
            if ph is None:
                continue
            ph_type = ph.get('type', '')
            texts = [t.text for t in sp.findall('.//a:t', ns) if t.text]
            if not texts:
                continue
            text = "".join(texts).strip()
            if not text:
                continue

            if ph_type == 'title' and not info['title']:
                info['title'] = text
            elif ph_type == 'subtitle' and not info['subtitle']:
                info['subtitle'] = text
            elif ph_type == 'ctrTitle' and not info['title']:
                info['title'] = text
            elif ph_type == 'body':
                # Heuristic: map first body line to author if empty
                if not info['author']:
                    info['author'] = text
                elif not info['institute']:
                    info['institute'] = text
                elif not info['date']:
                    info['date'] = text
    except Exception:
        pass

    return info

def convert_ppt_to_beamer_position(position, paper_width=12192000, paper_height=6858000):
    """Convert PowerPoint coordinates to LaTeX/Beamer positioning."""
    x_rel = position['x'] / paper_width
    y_rel = position['y'] / paper_height
    width_rel = position['width'] / paper_width
    height_rel = position['height'] / paper_height
    
    return {
        'x_rel': x_rel,
        'y_rel': y_rel,
        'width_rel': width_rel,
        'height_rel': height_rel
    }

def parse_presentation_xml(ppt_dir):
    """Parses presentation.xml for slide size."""
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    pres_xml = ppt_dir / "ppt" / "presentation.xml"
    size = {'width': 12192000, 'height': 6858000}
    if pres_xml.exists():
        try:
            tree = ET.parse(pres_xml)
            root = tree.getroot()
            sldSz = root.find('.//p:sldSz', ns)
            if sldSz is not None:
                size['width'] = int(sldSz.get('cx', 12192000))
                size['height'] = int(sldSz.get('cy', 6858000))
        except: pass
    return size

def is_full_slide_background(pos_rel, tol=0.02):
    """Heuristic: detect full-slide background images."""
    return (
        pos_rel['x_rel'] <= tol and
        pos_rel['y_rel'] <= tol and
        pos_rel['width_rel'] >= (1.0 - tol) and
        pos_rel['height_rel'] >= (1.0 - tol)
    )

def generate_main_tex(output_dir, slides_content, paper_width, paper_height, title_info):
    """Generates the main .tex file using the SWIFT_lecture_notes template structure."""
    filepath = output_dir / "overview_eng.tex"
    with open(filepath, 'w') as f:
        f.write(r"\documentclass{beamer}" + "\n")
        f.write(r"\input{../0-package.tex}" + "\n")
        f.write(r"\input{../0-macro.tex}" + "\n\n")
        
        f.write(rf"\author{{{escape_latex(title_info.get('author', ''))}}}" + "\n")
        f.write(rf"\title{{{escape_latex(title_info.get('title', ''))}}}" + "\n")
        f.write(rf"\subtitle{{{escape_latex(title_info.get('subtitle', ''))}}}" + "\n")
        f.write(rf"\institute{{{escape_latex(title_info.get('institute', ''))}}}" + "\n")
        f.write(rf"\date{{{escape_latex(title_info.get('date', ''))}}}" + "\n")
        f.write("\n")
        
        f.write(r"\begin{document}" + "\n\n")
        f.write(r"\kaishu" + "\n")
        f.write(r"\begin{frame}" + "\n")
        f.write(r"    \titlepage" + "\n")
        f.write(r"    \begin{figure}[htpb]" + "\n")
        f.write(r"        \begin{center}" + "\n")
        f.write(r"            \includegraphics[width=0.618\linewidth]{../pic/szu_logo.png}" + "\n")
        f.write(r"        \end{center}" + "\n")
        f.write(r"    \end{figure}" + "\n")
        f.write(r"\end{frame}" + "\n\n")
        
        if slides_content:
            for slide in slides_content:
                if str(slide['number']) == "1":
                    continue
                f.write(f"% Slide {slide['number']}\n")
                f.write(r"\begin{frame}" + f"{{{escape_latex(slide['title'])}}}\n")
                
                if slide['texts']:
                    f.write(r"  \begin{itemize}" + "\n")
                    for text in slide['texts']:
                        f.write(f"    \\item {escape_latex(text)}\n")
                    f.write(r"  \end{itemize}" + "\n")

                for img in slide['images']:
                    pos = convert_ppt_to_beamer_position(img['position'], paper_width, paper_height)
                    img_path = img['name']
                    if img_path.lower().endswith('.emf'):
                        img_path = img_path.replace('.emf', '.pdf')
                    if img_path.lower().endswith(('.tiff', '.tif')):
                        continue

                        img_path = f"fig/{img_path}"

                    if is_full_slide_background(pos):
                        continue

                    f.write(r"  \begin{center}" + "\n")
                    f.write(f"    \\includegraphics[width=0.85\\linewidth]{{{img_path}}}\n")
                    f.write(r"  \end{center}" + "\n")
                f.write(r"\end{frame}" + "\n\n")
        
        f.write(r"\end{document}" + "\n")

# --- Main Function ---

def write_shared_tex_files(shared_dir):
    """Write shared package/macro files to a common parent directory."""
    package_file = shared_dir / "0-package.tex"
    macro_file = shared_dir / "0-macro.tex"

    package_content = [
        r"\usepackage{ctex, hyperref}",
        r"\usepackage[T1]{fontenc}",
        r"\usepackage{latexsym,amsmath,xcolor,multicol,booktabs,calligra}",
        r"\usepackage{graphicx,pstricks,listings,stackengine}",
        r"\graphicspath{{fig/}}",
        r"\usepackage{tikz}",
        r"\usepackage{../szu_blue}",
        ""
    ]

    macro_content = [
        "% Put shared macros here",
        ""
    ]

    package_file.write_text("\n".join(package_content))
    macro_file.write_text("\n".join(macro_content))

def main():
    parser = argparse.ArgumentParser(description="Extract PPTX content into SWIFT Beamer template.")
    parser.add_argument("pptx_file", type=Path, help="Path to the input .pptx file")
    parser.add_argument("--output-dir", "-o", type=str, default="tex/overview_eng", help="Output directory")

    args = parser.parse_args()
    output_dir = Path(args.output_dir)

    if not args.pptx_file.is_file():
        print(f"Error: '{args.pptx_file}' not found.")
        sys.exit(1)

    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    # Copy template files from SWIFT_lecture_notes
    template_dir = Path("/Users/miranda/git/pptx2beamer/template")
    if template_dir.exists():
        for item in template_dir.iterdir():
            if item.name == "main.tex": continue
            if item.is_dir():
                if item.name == "pic":
                    shutil.copytree(item, output_dir.parent / item.name, dirs_exist_ok=True)
                else:
                    shutil.copytree(item, output_dir / item.name, dirs_exist_ok=True)
            else:
                if item.suffix == ".sty":
                    shutil.copy2(item, output_dir.parent / item.name)
                else:
                    shutil.copy2(item, output_dir / item.name)

    # Write shared package/macro files in parent directory
    write_shared_tex_files(output_dir.parent)

    # Ensure figure directory exists
    fig_dir = output_dir / "fig"
    fig_dir.mkdir(parents=True, exist_ok=True)

    # Process PowerPoint file
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        with zipfile.ZipFile(args.pptx_file, 'r') as zip_ref:
            zip_ref.extractall(temp_path)

        slide_size = parse_presentation_xml(temp_path)
        paper_width = slide_size['width']
        paper_height = slide_size['height']
        slides_content = parse_slides_for_content(temp_path)
        title_info = parse_title_page_info(temp_path)

        # Copy media files into fig/<filename>
        media_path = temp_path / "ppt" / "media"
        if media_path.exists():
            for media_file in media_path.iterdir():
                if media_file.is_file():
                    target_name = media_file.name
                    if target_name.lower().endswith('.jfif'):
                        target_name = target_name[:-5] + ".jpg"
                    shutil.copy2(media_file, fig_dir / target_name)

        # Generate final tex
        generate_main_tex(output_dir, slides_content, paper_width, paper_height, title_info)

    print(f"\n🎉 Content extracted into: {output_dir}/")
    print(f"Template files from SWIFT_lecture_notes copied.")

if __name__ == "__main__":
    main()
