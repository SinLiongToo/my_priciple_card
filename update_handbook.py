import re
import os
import sys
import json
import pptx

# Reconfigure stdout to use UTF-8
sys.stdout.reconfigure(encoding='utf-8')

def load_dotenv():
    """Load environment variables from .env file if it exists."""
    if os.path.exists(".env"):
        with open(".env", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ[k.strip()] = v.strip()

load_dotenv()

def parse_pptx_items(pptx_path):
    """Parse slide items from PPTX with the sub-item heuristic and handle duplicate numbers."""
    prs = pptx.Presentation(pptx_path)
    parsed_items = []
    
    for slide_idx, slide in enumerate(prs.slides, 1):
        slide_text_blocks = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if "☆" in text:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    p_text = paragraph.text.strip()
                    if p_text:
                        slide_text_blocks.append(p_text)
                        
        current_num = None
        current_lines = []
        
        def flush_item():
            nonlocal current_num, current_lines
            if current_num is not None and current_lines:
                text = " ".join(current_lines).strip()
                text = re.sub(r"\s+", " ", text)
                parsed_items.append({
                    "slide": slide_idx,
                    "num": current_num,
                    "text": text
                })
            current_lines = []
            current_num = None

        for block in slide_text_blocks:
            match = re.match(r"^(\d+)\s*\.\s*(.*)", block, re.DOTALL)
            if match:
                num = int(match.group(1))
                text = match.group(2)
                # Heuristic: if current_num exists and the new number is much smaller, it's a subitem
                if current_num is not None and (current_num - num > 20):
                    current_lines.append(block)
                else:
                    flush_item()
                    current_num = num
                    current_lines.append(text)
            else:
                if current_num is not None:
                    current_lines.append(block)
                else:
                    lines = block.split('\n')
                    for line in lines:
                        line = line.strip()
                        m = re.match(r"^(\d+)\s*\.\s*(.*)", line)
                        if m:
                            num = int(m.group(1))
                            t = m.group(2)
                            if current_num is not None and (current_num - num > 20):
                                current_lines.append(line)
                            else:
                                flush_item()
                                current_num = num
                                current_lines.append(t)
                        elif current_num is not None:
                            current_lines.append(line)
        flush_item()

    # Suffix duplicate numbers
    num_counts = {}
    for item in parsed_items:
        num = item['num']
        num_counts[num] = num_counts.get(num, 0) + 1
        
    seen_counts = {}
    for item in parsed_items:
        num = item['num']
        if num_counts[num] > 1:
            seen_counts[num] = seen_counts.get(num, 0) + 1
            item['key'] = f"{num}-{seen_counts[num]}"
        else:
            item['key'] = str(num)
            
    return parsed_items

def parse_existing_md(md_path):
    """Parse Part 1 (Principles 1-20) and other metadata from the existing markdown file."""
    if not os.path.exists(md_path):
        return "", []
        
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()
        
    parts = content.split("## 第二部分：簡報對照與逐條潤稿 (簡報版)")
    part1_content = parts[0]
    
    # Parse principles
    principle_pattern = re.compile(r"#### 原則 (\d+)：(.*?)\n(.*?)(?=\n#### 原則|\n##|\n---|\Z)", re.DOTALL)
    principles = []
    
    for match in principle_pattern.finditer(part1_content):
        num = int(match.group(1))
        title = match.group(2).strip()
        body = match.group(3).strip()
        
        mandarin = ""
        english = ""
        taiwanese = ""
        ref = ""
        
        m_match = re.search(r"\*\s+\*\*國語 \(Mandarin\)\*\*：(.*?)(?=\n\*|\n\[對照|\Z)", body)
        e_match = re.search(r"\*\s+\*\*English\*\*：(.*?)(?=\n\*|\n\[對照|\Z)", body)
        t_match = re.search(r"\*\s+\*\*台語 \(Taiwanese\)\*\*：(.*?)(?=\n\*|\n\[對照|\Z)", body)
        ref_match = re.search(r"`\[對照原項目：(.*?)\]`", body)
        
        if m_match: mandarin = m_match.group(1).strip()
        if e_match: english = e_match.group(1).strip()
        if t_match: taiwanese = t_match.group(1).strip()
        if ref_match: ref = ref_match.group(1).strip()
        
        # Clean ref brackets if any
        ref = ref.replace("[對照原項目：", "").replace("]", "").strip()
        
        principles.append({
            "num": num,
            "title": title,
            "mandarin": mandarin,
            "english": english,
            "taiwanese": taiwanese,
            "ref": ref
        })
        
    return part1_content, principles

def translate_via_gemini(original_text):
    """Translate original English text into Mandarin, English, and Taiwanese using Gemini API."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return None
        
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        
        prompt = f"""
You are an expert translator and copywriter specializing in Taiwanese Hokkien (台語漢字) and professional business context.
Please translate and polish the following workplace principle:
"{original_text}"

You must return a JSON object with exactly the following keys:
- "english": Polished, professional English version.
- "mandarin": Polished, natural Traditional Chinese (Mandarin/國語) version suitable for business context.
- "taiwanese": Idiomatic Taiwanese Hokkien (台語漢字) version written in standard Traditional Chinese characters representing spoken Taiwanese, capturing the business wisdom naturally (e.g. use standard terms like "鬥相共", "省氣力", "做代誌", etc.).

Return ONLY the raw JSON string without markdown blocks or explanation.
"""
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        
        # Parse JSON from response
        res_text = response.text.strip()
        # strip markdown formatting if the model returned it
        if res_text.startswith("```"):
            res_text = re.sub(r"^```(?:json)?\n", "", res_text)
            res_text = re.sub(r"\n```$", "", res_text)
            
        data = json.loads(res_text.strip())
        return data
    except Exception as e:
        print(f"Error calling Gemini API: {e}")
        return None

def generate_odt_handbook(principles, final_items, odt_path):
    """Generate a well-formatted ODT document for the handbook ready for printing/book publishing."""
    from odf.opendocument import OpenDocumentText
    from odf.style import Style, TextProperties, ParagraphProperties, PageLayout, PageLayoutProperties, MasterPage
    from odf.text import H, P, Span

    print("\nStep 7: Generating ODT handbook for publishing...")
    doc = OpenDocumentText()
    
    # ------------------ DEFINE PAGE LAYOUT (A5) ------------------
    page_layout = PageLayout(name="A5Layout")
    page_layout.addElement(PageLayoutProperties(
        pagewidth="14.8cm",
        pageheight="21.0cm",
        marginleft="1.5cm",
        marginright="1.5cm",
        margintop="1.5cm",
        marginbottom="1.5cm"
    ))
    doc.automaticstyles.addElement(page_layout)
    
    master_page = MasterPage(name="Standard", pagelayoutname=page_layout)
    doc.masterstyles.addElement(master_page)
    
    # ------------------ DEFINE STYLES ------------------
    # Title style
    title_style = Style(name="BookTitle", family="paragraph")
    title_style.addElement(TextProperties(fontsize="26pt", fontweight="bold", fontfamily="Microsoft JhengHei", color="#1e293b"))
    title_style.addElement(ParagraphProperties(textalign="center", margintop="2cm", marginbottom="0.5cm"))
    doc.styles.addElement(title_style)
    
    # Subtitle style
    subtitle_style = Style(name="BookSubtitle", family="paragraph")
    subtitle_style.addElement(TextProperties(fontsize="14pt", fontfamily="Microsoft JhengHei", color="#4f46e5"))
    subtitle_style.addElement(ParagraphProperties(textalign="center", margintop="0.2cm", marginbottom="3cm"))
    doc.styles.addElement(subtitle_style)
    
    # Meta style (author, date etc.)
    meta_style = Style(name="BookMeta", family="paragraph")
    meta_style.addElement(TextProperties(fontsize="10pt", fontfamily="Microsoft JhengHei", color="#64748b"))
    meta_style.addElement(ParagraphProperties(textalign="center", margintop="5cm", marginbottom="1cm"))
    doc.styles.addElement(meta_style)
    
    # Chapter Heading style (H1 with Page Break, Centered for Divider Page)
    chapter_style = Style(name="BookChapter", family="paragraph")
    chapter_style.addElement(TextProperties(fontsize="20pt", fontweight="bold", fontfamily="Microsoft JhengHei", color="#4f46e5"))
    chapter_style.addElement(ParagraphProperties(breakbefore="page", textalign="center", margintop="4.5cm", marginbottom="1cm"))
    doc.styles.addElement(chapter_style)
    
    # Principle Heading style (H2 with Page Break)
    principle_title_style = Style(name="BookPrincipleTitle", family="paragraph")
    principle_title_style.addElement(TextProperties(fontsize="14pt", fontweight="bold", fontfamily="Microsoft JhengHei", color="#1e293b"))
    principle_title_style.addElement(ParagraphProperties(breakbefore="page", margintop="0.8cm", marginbottom="0.4cm", keepwithnext="true"))
    doc.styles.addElement(principle_title_style)
    
    # Principle Translation Body style
    principle_body_style = Style(name="BookPrincipleBody", family="paragraph")
    principle_body_style.addElement(TextProperties(fontsize="10.5pt", fontfamily="Microsoft JhengHei", color="#334155"))
    principle_body_style.addElement(ParagraphProperties(margintop="0.1cm", marginbottom="0.1cm", lineheight="150%"))
    doc.styles.addElement(principle_body_style)
    
    # Bold inline label style
    bold_label_style = Style(name="BoldLabel", family="text")
    bold_label_style.addElement(TextProperties(fontweight="bold", color="#1e293b"))
    doc.styles.addElement(bold_label_style)
    
    # Slide Item Heading style (H3 without page break)
    slide_title_style = Style(name="BookSlideTitle", family="paragraph")
    slide_title_style.addElement(TextProperties(fontsize="11pt", fontweight="bold", fontfamily="Microsoft JhengHei", color="#0f766e")) # Teal
    slide_title_style.addElement(ParagraphProperties(margintop="0.6cm", marginbottom="0.2cm", keepwithnext="true"))
    doc.styles.addElement(slide_title_style)
    
    # Slide Item Heading style with page break (H3 with page break)
    slide_title_style_with_break = Style(name="BookSlideTitleWithBreak", family="paragraph")
    slide_title_style_with_break.addElement(TextProperties(fontsize="11pt", fontweight="bold", fontfamily="Microsoft JhengHei", color="#0f766e")) # Teal
    slide_title_style_with_break.addElement(ParagraphProperties(breakbefore="page", margintop="0.6cm", marginbottom="0.2cm", keepwithnext="true"))
    doc.styles.addElement(slide_title_style_with_break)
    
    # Slide Item Original style (Monospace)
    slide_orig_style = Style(name="BookSlideOrig", family="paragraph")
    slide_orig_style.addElement(TextProperties(fontsize="9.5pt", fontfamily="Courier New", color="#475569", fontstyle="italic"))
    slide_orig_style.addElement(ParagraphProperties(margintop="0.1cm", marginbottom="0.1cm", marginleft="0.5cm"))
    doc.styles.addElement(slide_orig_style)
    
    # Slide Item Translation style
    slide_trans_style = Style(name="BookSlideTrans", family="paragraph")
    slide_trans_style.addElement(TextProperties(fontsize="9.5pt", fontfamily="Microsoft JhengHei", color="#334155"))
    slide_trans_style.addElement(ParagraphProperties(margintop="0.1cm", marginbottom="0.1cm", marginleft="0.5cm", lineheight="140%"))
    doc.styles.addElement(slide_trans_style)
    
    # ------------------ BUILD COVER PAGE ------------------
    p = P(stylename=title_style)
    p.addText("工作的管見")
    doc.text.addElement(p)
    
    p = P(stylename=subtitle_style)
    p.addText("職場生存與成長的避坑指南 (三語對照手冊)")
    doc.text.addElement(p)
    
    p = P(stylename=meta_style)
    p.addText("精裝出書版面格式\n© 2026 工作的管見")
    doc.text.addElement(p)
    
    # ------------------ GENERATE CHAPTERS & PRINCIPLES ------------------
    def get_chapter_name(num):
        if num <= 4: return '第一篇：效能與工具（生產力管理）'
        if num <= 8: return '第二篇：思考與決策（思維模型與問題解決）'
        if num <= 11: return '第三篇：溝通與協作（職場對話術）'
        if num <= 14: return '第四篇：職場生存與防線（自我保護與邊界）'
        if num <= 17: return '第五篇：職涯與個人成長（價值提升與升遷）'
        return '第六篇：能量、韌性與生活平衡（心智與身心管理）'

    current_chapter = None
    
    for p_item in principles:
        num = p_item['num']
        chap_name = get_chapter_name(num)
        
        # Add Chapter Heading if changed
        if chap_name != current_chapter:
            current_chapter = chap_name
            h_chap = H(outlinelevel=1, stylename=chapter_style)
            h_chap.addText(current_chapter)
            doc.text.addElement(h_chap)
            
        # Add Principle Title (Always starts on new page)
        h_pr = H(outlinelevel=2, stylename=principle_title_style)
        h_pr.addText(f"原則 {num}：{p_item['title']}")
        doc.text.addElement(h_pr)
        
        # Add Principle Content
        def add_translation_row(doc, label, text):
            p = P(stylename=principle_body_style)
            span = Span(stylename=bold_label_style)
            span.addText(label + "：")
            p.addElement(span)
            p.addText(text)
            doc.text.addElement(p)
            
        add_translation_row(doc, "國語 (Mandarin)", p_item['mandarin'])
        add_translation_row(doc, "English", p_item['english'])
        add_translation_row(doc, "台語 (Taiwanese)", p_item['taiwanese'])
        
        # Add Slide Items corresponding to this Principle
        refs = [r.strip() for r in p_item['ref'].split(',')]
        matched_items = []
        for item in final_items:
            item_key = item['key']
            base_key = item_key.split('-')[0]
            if item_key in refs or base_key in refs:
                matched_items.append(item)
                
        if matched_items:
            for idx, item in enumerate(matched_items):
                # Page break before first item (so Principle is on separate page)
                # and after every 2 items (so max 2 items per page)
                h_style = slide_title_style_with_break if idx % 2 == 0 else slide_title_style
                
                # Add Slide Item Title
                h_slide = H(outlinelevel=3, stylename=h_style)
                h_slide.addText(f"No. {item['key']} (Slide {item['slide']})")
                doc.text.addElement(h_slide)
                
                # Original Text
                p_orig = P(stylename=slide_orig_style)
                span = Span(stylename=bold_label_style)
                span.addText("原文：")
                p_orig.addElement(span)
                p_orig.addText(item['original'])
                doc.text.addElement(p_orig)
                
                # Mandarin
                p_mand = P(stylename=slide_trans_style)
                span = Span(stylename=bold_label_style)
                span.addText("優化中文：")
                p_mand.addElement(span)
                p_mand.addText(item['mandarin'])
                doc.text.addElement(p_mand)
                
                # English
                p_eng = P(stylename=slide_trans_style)
                span = Span(stylename=bold_label_style)
                span.addText("優化英文：")
                p_eng.addElement(span)
                p_eng.addText(item['english'])
                doc.text.addElement(p_eng)
                
                # Taiwanese
                p_taiw = P(stylename=slide_trans_style)
                span = Span(stylename=bold_label_style)
                span.addText("優化台文：")
                p_taiw.addElement(span)
                p_taiw.addText(item['taiwanese'])
                doc.text.addElement(p_taiw)
                
    doc.save(odt_path)
    print(f"Saved regenerated ODT handbook: {odt_path}")

def main():
    pptx_path = "工作原則-書.pptx.pptx"
    md_path = "工作原則整理與潤稿.md"
    db_path = "translation_db.json"
    
    print("Step 1: Parsing PowerPoint slides...")
    pptx_items = parse_pptx_items(pptx_path)
    print(f"Found {len(pptx_items)} items in PPTX.")
    
    print("\nStep 2: Loading translation database...")
    db = {}
    if os.path.exists(db_path):
        with open(db_path, "r", encoding="utf-8") as f:
            db = json.load(f)
    print(f"Loaded {len(db)} entries from database.")
    
    # Track missing or modified items to translate
    missing_items = []
    
    print("\nStep 3: Matching items and verifying translations...")
    final_items = []
    translated_new_count = 0
    
    for item in pptx_items:
        key = item['key']
        original = item['text']
        
        # Check database by original text match
        db_match = db.get(original)
        
        # If not found, try to look up with minor space differences
        if not db_match:
            original_clean = re.sub(r'\s+', '', original)
            for db_orig, db_val in db.items():
                if re.sub(r'\s+', '', db_orig) == original_clean:
                    db_match = db_val
                    break
                    
        if db_match:
            final_items.append({
                "key": key,
                "slide": item['slide'],
                "original": original,
                "english": db_match['english'],
                "mandarin": db_match['mandarin'],
                "taiwanese": db_match['taiwanese']
            })
        else:
            # We need to translate this
            print(f"  Missing translation for Key {key}: '{original[:60]}...'")
            missing_items.append((original, key, item['slide']))

    # Perform translations if API key is present
    api_key = os.environ.get("GEMINI_API_KEY")
    if missing_items:
        if api_key:
            print(f"\nFound {len(missing_items)} items requiring translation. Running API translations...")
            for idx, (orig, key, slide) in enumerate(missing_items, 1):
                print(f"  ({idx}/{len(missing_items)}) Translating Key {key}...")
                translated = translate_via_gemini(orig)
                if translated:
                    # Save to db
                    db[orig] = {
                        "english": translated['english'],
                        "mandarin": translated['mandarin'],
                        "taiwanese": translated['taiwanese'],
                        "key": key
                    }
                    final_items.append({
                        "key": key,
                        "slide": slide,
                        "original": orig,
                        "english": translated['english'],
                        "mandarin": translated['mandarin'],
                        "taiwanese": translated['taiwanese']
                    })
                    translated_new_count += 1
                else:
                    # Fallback placeholders
                    final_items.append({
                        "key": key,
                        "slide": slide,
                        "original": orig,
                        "english": f"[TBD] {orig}",
                        "mandarin": f"[TBD] 待翻譯 - 原文: {orig}",
                        "taiwanese": f"[TBD] 待翻譯 - 原文: {orig}"
                    })
            # Save updated DB
            with open(db_path, "w", encoding="utf-8") as f:
                json.dump(db, f, ensure_ascii=False, indent=2)
            print(f"API translation complete. Added {translated_new_count} translations to database.")
        else:
            print(f"\n[WARNING] Found {len(missing_items)} items requiring translation, but GEMINI_API_KEY is not set.")
            print("Placeholders will be generated. Please set GEMINI_API_KEY in `.env` to automatically translate them.")
            for orig, key, slide in missing_items:
                final_items.append({
                    "key": key,
                    "slide": slide,
                    "original": orig,
                    "english": f"[TBD] {orig}",
                    "mandarin": f"[TBD] 待翻譯 - 原文: {orig}",
                    "taiwanese": f"[TBD] 待翻譯 - 原文: {orig}"
                })

    # Sort final items by key number
    def get_key_sort(x):
        k = x['key']
        # Split number and suffix
        parts = re.split(r'(\d+)', k)
        return [int(s) if s.isdigit() else s for s in parts]
        
    final_items.sort(key=get_key_sort)
    
    print("\nStep 4: Parsing Part 1 (Principles) from existing markdown...")
    part1_content, principles = parse_existing_md(md_path)
    
    if not part1_content:
        # Fallback default title/toc if file does not exist
        part1_content = "# 工作的管見：職場生存與成長的避坑指南\n## 工作原則整理與潤稿 (三語對照版)\n\n"
        
    print("\nStep 5: Regenerating Markdown handbook...")
    # Write Part 1 and regenerated Part 2
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(part1_content.strip() + "\n\n---\n\n## 第二部分：簡報對照與逐條潤稿 (簡報版)\n\n")
        f.write(f"> [!NOTE]\n> 本部分為逐頁對照簡報 (PPTX) 的 {len(final_items)} 條原始筆記進行優化與翻譯的內容，方便您在更新投影片時直接複製貼上使用。\n\n")
        
        current_slide = 0
        for item in final_items:
            slide = item['slide']
            if slide != current_slide:
                current_slide = slide
                # Find min/max keys on this slide
                slide_items = [x for x in final_items if x['slide'] == slide]
                min_key = slide_items[0]['key']
                max_key = slide_items[-1]['key']
                f.write(f"### Slide {slide} (原則 {min_key} - {max_key})\n\n")
                
            # Find matched principles
            matched_principles = []
            item_key = item['key']
            base_key = item_key.split('-')[0]
            for p in principles:
                refs = [r.strip() for r in p['ref'].split(',')]
                if item_key in refs or base_key in refs:
                    matched_principles.append(p['num'])
            
            if matched_principles:
                principles_str = ", ".join(f"原則 {num}" for num in matched_principles)
                f.write(f"#### **No. {item['key']}** ({principles_str})\n")
            else:
                f.write(f"#### **No. {item['key']}**\n")
            f.write(f"*   **原文**：`{item['original']}`\n")
            f.write(f"*   **優化英文**：{item['english']}\n")
            f.write(f"*   **優化中文**：{item['mandarin']}\n")
            f.write(f"*   **優化台文**：{item['taiwanese']}\n\n")
            
    print(f"Saved regenerated handbook: {md_path}")
    
    print("\nStep 6: Generating HTML card deck web application...")
    
    # We serialize slides and principles as JSON variables in HTML
    principles_json = json.dumps(principles, ensure_ascii=False)
    slides_json = json.dumps(final_items, ensure_ascii=False)

    slide_groups = {}
    for item in final_items:
        s = item['slide']
        if s not in slide_groups: slide_groups[s] = []
        slide_groups[s].append(item['key'])
    slide_buttons = ['<button class="filter-tag active" onclick="filterSlides(\'all\')">全部 Slide</button>']
    for s in sorted(slide_groups.keys()):
        min_k = slide_groups[s][0]
        max_k = slide_groups[s][-1]
        slide_buttons.append(f'<button class="filter-tag" onclick="filterSlides(\'{s}\')">Slide {s} ({min_k}-{max_k})</button>')
    slides_filter_html = "\n                ".join(slide_buttons)
    
    # HTML template with embedded styling and logic
    html_content = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>工作的管見 - 工作原則卡牌</title>
    <meta name="description" content="工作的管見：職場生存與成長的避坑指南 - 三語對照數位原則卡牌與隨機閱讀器">
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&family=Noto+Sans+TC:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary: #6366f1;
            --primary-hover: #4f46e5;
            --bg-dark: #0f172a;
            --bg-card: rgba(30, 41, 59, 0.7);
            --bg-glass: rgba(15, 23, 42, 0.65);
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border: rgba(255, 255, 255, 0.08);
            --border-hover: rgba(99, 102, 241, 0.4);
            --glow: rgba(99, 102, 241, 0.15);
        }}

        .light-mode {{
            --primary: #4f46e5;
            --primary-hover: #3730a3;
            --bg-dark: #f8fafc;
            --bg-card: rgba(255, 255, 255, 0.85);
            --bg-glass: rgba(241, 245, 249, 0.75);
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border: rgba(0, 0, 0, 0.08);
            --border-hover: rgba(79, 70, 229, 0.4);
            --glow: rgba(79, 70, 229, 0.1);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            transition: background-color 0.3s, border-color 0.3s;
        }}

        body {{
            font-family: 'Outfit', 'Noto Sans TC', sans-serif;
            background-color: var(--bg-dark);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            overflow-x: hidden;
            background-image: 
                radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.12) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(139, 92, 246, 0.1) 0px, transparent 50%);
            background-attachment: fixed;
        }}

        /* Header block */
        header {{
            padding: 2.5rem 1.5rem 1.5rem;
            text-align: center;
            border-bottom: 1px solid var(--border);
            background: var(--bg-glass);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            position: sticky;
            top: 0;
            z-index: 100;
        }}

        header h1 {{
            font-size: 2.2rem;
            font-weight: 800;
            letter-spacing: -0.025em;
            background: linear-gradient(135deg, #a5b4fc 0%, #6366f1 50%, #818cf8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }}

        .light-mode header h1 {{
            background: linear-gradient(135deg, #4f46e5 0%, #312e81 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        header p {{
            color: var(--text-muted);
            font-size: 1rem;
            font-weight: 300;
            max-width: 600px;
            margin: 0 auto 1.5rem;
        }}

        /* Navigation and tools */
        .controls-container {{
            max-width: 1200px;
            width: 100%;
            margin: 0 auto;
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            align-items: center;
            justify-content: space-between;
        }}

        .tabs {{
            display: flex;
            background: rgba(0, 0, 0, 0.2);
            padding: 0.25rem;
            border-radius: 99px;
            border: 1px solid var(--border);
        }}

        .light-mode .tabs {{
            background: rgba(0, 0, 0, 0.05);
        }}

        .tab-btn {{
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 0.6rem 1.4rem;
            font-size: 0.95rem;
            font-weight: 600;
            border-radius: 99px;
            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        .tab-btn.active {{
            background: var(--primary);
            color: #ffffff;
            box-shadow: 0 4px 12px var(--glow);
        }}

        .action-tools {{
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }}

        .search-bar {{
            position: relative;
            min-width: 260px;
        }}

        .search-input {{
            width: 100%;
            padding: 0.6rem 1rem 0.6rem 2.5rem;
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid var(--border);
            border-radius: 99px;
            color: var(--text-main);
            font-size: 0.9rem;
            outline: none;
            transition: all 0.2s;
        }}

        .light-mode .search-input {{
            background: rgba(255, 255, 255, 0.9);
        }}

        .search-input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 3px var(--glow);
        }}

        .search-icon {{
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            pointer-events: none;
            font-size: 0.9rem;
        }}

        .icon-btn {{
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid var(--border);
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            color: var(--text-main);
            outline: none;
            font-size: 1.1rem;
        }}

        .light-mode .icon-btn {{
            background: rgba(255, 255, 255, 0.9);
        }}

        .icon-btn:hover {{
            border-color: var(--primary);
            color: var(--primary);
        }}

        /* Main layouts */
        main {{
            flex: 1;
            max-width: 1200px;
            width: 100%;
            margin: 2rem auto;
            padding: 0 1.5rem;
        }}

        .section-view {{
            display: none;
        }}

        .section-view.active {{
            display: block;
            animation: fadeIn 0.4s ease-out;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        /* Filters */
        .filter-tags {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-bottom: 1.5rem;
        }}

        .filter-tag {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border);
            padding: 0.4rem 1rem;
            border-radius: 99px;
            font-size: 0.85rem;
            cursor: pointer;
            color: var(--text-muted);
            font-weight: 500;
        }}

        .light-mode .filter-tag {{
            background: rgba(0, 0, 0, 0.02);
        }}

        .filter-tag:hover, .filter-tag.active {{
            border-color: var(--primary);
            color: var(--primary);
            background: rgba(99, 102, 241, 0.05);
        }}

        /* Card grid */
        .card-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
            gap: 1.5rem;
        }}

        .principle-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 1.5rem;
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
            display: flex;
            flex-direction: column;
            position: relative;
            overflow: hidden;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        .principle-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, var(--primary) 0%, #818cf8 100%);
            opacity: 0;
            transition: opacity 0.3s;
        }}

        .principle-card:hover {{
            transform: translateY(-5px);
            border-color: var(--border-hover);
            box-shadow: 0 12px 30px rgba(99, 102, 241, 0.08);
        }}

        .principle-card:hover::before {{
            opacity: 1;
        }}

        .card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
        }}

        .card-num {{
            font-size: 0.8rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--primary);
            background: rgba(99, 102, 241, 0.1);
            padding: 0.25rem 0.6rem;
            border-radius: 4px;
        }}

        .card-principle-badge {{
            font-size: 0.8rem;
            font-weight: 800;
            letter-spacing: 0.05em;
            color: #10b981;
            background: rgba(16, 185, 129, 0.1);
            padding: 0.25rem 0.6rem;
            border-radius: 4px;
            display: inline-block;
        }}

        .light-mode .card-principle-badge {{
            color: #059669;
            background: rgba(5, 150, 105, 0.1);
        }}

        .ref-items-container {{
            margin-top: 1.2rem;
            padding-top: 0.8rem;
            border-top: 1px dashed var(--border);
        }}

        .ref-items-label {{
            font-size: 0.78rem;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 0.6rem;
            display: block;
        }}

        .ref-chips-grid {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.4rem;
        }}

        .ref-item-chip {{
            font-size: 0.75rem;
            font-weight: 600;
            font-family: inherit;
            color: var(--primary);
            background: rgba(99, 102, 241, 0.12);
            border: 1px solid rgba(99, 102, 241, 0.25);
            padding: 0.25rem 0.55rem;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
        }}

        .light-mode .ref-item-chip {{
            background: rgba(79, 70, 229, 0.08);
            border-color: rgba(79, 70, 229, 0.2);
            color: var(--primary);
        }}

        .ref-item-chip:hover {{
            background: var(--primary);
            color: #ffffff;
            border-color: var(--primary);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.35);
        }}

        .highlight-pulse {{
            animation: pulseGlow 2.5s ease-in-out infinite;
            border-color: var(--primary) !important;
            z-index: 10;
        }}

        @keyframes pulseGlow {{
            0% {{ box-shadow: 0 0 0 0 rgba(99, 102, 241, 0.8); transform: scale(1); }}
            50% {{ box-shadow: 0 0 30px 10px rgba(99, 102, 241, 0.6); transform: scale(1.03); }}
            100% {{ box-shadow: 0 0 0 0 rgba(99, 102, 241, 0); transform: scale(1); }}
        }}

        .card-meta {{
            font-size: 0.8rem;
            color: var(--text-muted);
        }}

        .card-title {{
            font-size: 1.2rem;
            font-weight: 700;
            margin-bottom: 1rem;
            line-height: 1.3;
        }}

        .lang-section {{
            margin-bottom: 0.75rem;
            font-size: 0.95rem;
            line-height: 1.5;
        }}

        .lang-section:last-child {{
            margin-bottom: 0;
        }}

        .lang-label {{
            font-size: 0.75rem;
            font-weight: 800;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            display: block;
            margin-bottom: 0.2rem;
        }}

        .lang-txt {{
            color: var(--text-main);
        }}

        .original-eng {{
            font-style: italic;
            color: var(--text-muted);
            border-left: 2px solid var(--border);
            padding-left: 0.75rem;
            margin-bottom: 1rem;
            font-size: 0.9rem;
        }}

        /* Random Card drawer view */
        .random-view-container {{
            max-width: 580px;
            margin: 2rem auto;
            display: flex;
            flex-direction: column;
            align-items: center;
        }}

        /* 3D card flip setup */
        .flip-card-container {{
            perspective: 1200px;
            width: 100%;
            height: 480px;
            cursor: pointer;
            margin-bottom: 2rem;
        }}

        .flip-card {{
            position: relative;
            width: 100%;
            height: 100%;
            transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
            transform-style: preserve-3d;
        }}

        .flip-card-container.flipped .flip-card {{
            transform: rotateY(180deg);
        }}

        .card-face {{
            position: absolute;
            width: 100%;
            height: 100%;
            backface-visibility: hidden;
            -webkit-backface-visibility: hidden;
            border-radius: 24px;
            border: 1px solid var(--border);
            padding: 2rem;
            box-shadow: 0 15px 35px rgba(0,0,0,0.2);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            background: var(--bg-card);
            backdrop-filter: blur(25px);
            -webkit-backdrop-filter: blur(25px);
        }}

        .card-face-front {{
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0.9) 100%);
            border-color: rgba(99, 102, 241, 0.25);
        }}

        .light-mode .card-face-front {{
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.9) 0%, rgba(241, 245, 249, 0.95) 100%);
        }}

        .card-face-back {{
            transform: rotateY(180deg);
            background: linear-gradient(135deg, rgba(15, 23, 42, 0.95) 0%, rgba(30, 41, 59, 0.8) 100%);
            border-color: rgba(139, 92, 246, 0.25);
        }}

        .light-mode .card-face-back {{
            background: linear-gradient(135deg, rgba(241, 245, 249, 0.95) 0%, rgba(255, 255, 255, 0.9) 100%);
        }}

        .front-body {{
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            flex: 1;
            text-align: center;
        }}

        .front-title {{
            font-size: 1.5rem;
            font-weight: 600;
            line-height: 1.5;
            color: var(--text-main);
        }}

        .back-body {{
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 1.25rem;
            flex: 1;
        }}

        .flip-hint {{
            font-size: 0.8rem;
            color: var(--text-muted);
            text-align: center;
            letter-spacing: 0.05em;
            text-transform: uppercase;
        }}

        .draw-btn {{
            background: linear-gradient(90deg, var(--primary) 0%, #818cf8 100%);
            color: white;
            border: none;
            padding: 0.9rem 2.5rem;
            font-size: 1.05rem;
            font-weight: 700;
            border-radius: 99px;
            cursor: pointer;
            box-shadow: 0 4px 20px var(--glow);
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }}

        .draw-btn:hover {{
            transform: scale(1.05);
            box-shadow: 0 8px 25px var(--glow);
        }}

        .draw-btn:active {{
            transform: scale(0.98);
        }}

        /* Particle effect */
        .sparkle-particles {{
            position: absolute;
            pointer-events: none;
            width: 10px;
            height: 10px;
            background: var(--primary);
            border-radius: 50%;
            opacity: 0;
        }}

        /* Footer block */
        footer {{
            text-align: center;
            padding: 2.5rem 1.5rem;
            border-top: 1px solid var(--border);
            color: var(--text-muted);
            font-size: 0.85rem;
            margin-top: auto;
        }}

        /* Visibility Controls & Toggles */
        .display-toggles-bar {{
            max-width: 1200px;
            width: 100%;
            margin: 1.2rem auto 0;
            padding-top: 0.9rem;
            border-top: 1px dashed var(--border);
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
            font-size: 0.85rem;
        }}

        .toggle-group-left {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 0.5rem;
        }}

        .toggle-group-label {{
            color: var(--text-muted);
            font-weight: 600;
            font-size: 0.82rem;
            display: flex;
            align-items: center;
            gap: 0.3rem;
            margin-right: 0.25rem;
        }}

        .toggle-chips {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.4rem;
            align-items: center;
        }}

        .toggle-chip {{
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            padding: 0.35rem 0.8rem;
            border-radius: 99px;
            border: 1px solid var(--border);
            background: rgba(0, 0, 0, 0.2);
            color: var(--text-muted);
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            user-select: none;
        }}

        .light-mode .toggle-chip {{
            background: rgba(0, 0, 0, 0.04);
        }}

        .chip-dot {{
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #64748b;
            transition: all 0.2s;
        }}

        .toggle-chip.active {{
            border-color: var(--primary);
            color: var(--text-main);
            background: rgba(99, 102, 241, 0.15);
        }}

        .light-mode .toggle-chip.active {{
            background: rgba(79, 70, 229, 0.1);
            color: var(--primary);
        }}

        .toggle-chip.active .chip-dot {{
            background: #10b981;
            box-shadow: 0 0 6px rgba(16, 185, 129, 0.6);
        }}

        .toggle-chip:hover {{
            border-color: var(--primary-hover);
            transform: translateY(-1px);
        }}

        .preset-links {{
            display: flex;
            align-items: center;
            gap: 0.4rem;
            color: var(--text-muted);
            font-size: 0.8rem;
        }}

        .preset-btn {{
            cursor: pointer;
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            font-size: 0.78rem;
            font-weight: 600;
            transition: all 0.2s;
        }}

        .light-mode .preset-btn {{
            background: rgba(0, 0, 0, 0.03);
        }}

        .preset-btn:hover {{
            color: var(--primary);
            border-color: var(--primary);
            background: rgba(99, 102, 241, 0.1);
        }}

        .front-placeholder {{
            display: none;
            font-size: 1.05rem;
            color: var(--text-muted);
            font-style: italic;
            padding: 1.5rem;
            text-align: center;
            line-height: 1.6;
        }}

        /* Dynamic field hiding rules */
        body.hide-original .original-eng,
        body.hide-original #randomFrontOriginal {{
            display: none !important;
        }}

        body.hide-original .front-placeholder {{
            display: block !important;
        }}

        body.hide-mandarin .lang-mandarin,
        body.hide-mandarin #randomBackMandarinSection {{
            display: none !important;
        }}

        body.hide-english .lang-english,
        body.hide-english #randomBackEnglishSection {{
            display: none !important;
        }}

        body.hide-taiwanese .lang-taiwanese,
        body.hide-taiwanese #randomBackTaiwaneseSection {{
            display: none !important;
        }}

        body.hide-principles .card-principle-badge,
        body.hide-principles .ref-items-container,
        body.hide-principles #randomFrontPrinciples,
        body.hide-principles #randomBackPrinciples {{
            display: none !important;
        }}

        .all-hidden-notice {{
            display: none;
            padding: 1.2rem;
            margin: 0.5rem 0;
            text-align: center;
            color: var(--text-muted);
            font-size: 0.85rem;
            border: 1px dashed var(--border);
            border-radius: 8px;
            background: rgba(0, 0, 0, 0.1);
        }}

        body.hide-mandarin.hide-english.hide-taiwanese #view-principles .all-hidden-notice,
        body.hide-mandarin.hide-english.hide-taiwanese #randomCardContainer .back-body .all-hidden-notice {{
            display: block !important;
        }}

        body.hide-original.hide-mandarin.hide-english.hide-taiwanese #view-slides .all-hidden-notice {{
            display: block !important;
        }}

        /* RWD Responsive Design */
        @media (max-width: 992px) {{
            .card-grid {{
                grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
                gap: 1.2rem;
            }}
        }}

        @media (max-width: 768px) {{
            header {{
                padding: 1.25rem 1rem 1rem;
            }}
            header h1 {{
                font-size: 1.65rem;
            }}
            header p {{
                font-size: 0.88rem;
                margin-bottom: 1rem;
            }}
            .controls-container {{
                flex-direction: column;
                align-items: stretch;
                gap: 0.75rem;
            }}
            .action-tools {{
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 0.6rem;
                width: 100%;
            }}
            .tabs {{
                width: 100%;
                display: flex;
                overflow-x: auto;
                white-space: nowrap;
                justify-content: space-between;
                -webkit-overflow-scrolling: touch;
            }}
            .tab-btn {{
                padding: 0.5rem 0.85rem;
                font-size: 0.85rem;
                flex: 1;
                text-align: center;
                white-space: nowrap;
            }}
            .search-bar {{
                flex: 1;
                min-width: 180px;
            }}
            .filter-tags {{
                overflow-x: auto;
                flex-wrap: nowrap;
                white-space: nowrap;
                padding-bottom: 0.4rem;
                margin-bottom: 1rem;
                -webkit-overflow-scrolling: touch;
                scrollbar-width: none;
            }}
            .filter-tags::-webkit-scrollbar {{
                display: none;
            }}
            .filter-tag {{
                flex-shrink: 0;
            }}
            .card-grid {{
                grid-template-columns: 1fr;
                gap: 1rem;
            }}
            .principle-card {{
                padding: 1.25rem 1.1rem;
            }}
            .card-title {{
                font-size: 1.1rem;
            }}
            main {{
                margin: 1.2rem auto;
                padding: 0 1rem;
            }}
            .display-toggles-bar {{
                flex-direction: column;
                align-items: flex-start;
                gap: 0.65rem;
                padding-top: 0.75rem;
            }}
            .toggle-group-left {{
                width: 100%;
            }}
            .toggle-chips {{
                width: 100%;
                overflow-x: auto;
                flex-wrap: nowrap;
                white-space: nowrap;
                padding-bottom: 0.25rem;
                -webkit-overflow-scrolling: touch;
                scrollbar-width: none;
            }}
            .toggle-chips::-webkit-scrollbar {{
                display: none;
            }}
            .toggle-chip {{
                flex-shrink: 0;
            }}
            .preset-links {{
                width: 100%;
                justify-content: flex-start;
                margin-top: 0.15rem;
            }}
            /* Random card mobile optimizations */
            .flip-card-container {{
                height: 480px;
                max-width: 100%;
            }}
            .card-face {{
                padding: 1.3rem 1.1rem;
            }}
            .front-title {{
                font-size: 1.2rem;
                line-height: 1.5;
            }}
            .back-body {{
                overflow-y: auto;
                padding-right: 0.2rem;
                gap: 0.85rem;
            }}
            .lang-section {{
                font-size: 0.9rem;
            }}
            .draw-btn {{
                width: 100%;
                justify-content: center;
                padding: 0.85rem 1.5rem;
            }}
        }}

        @media (max-width: 480px) {{
            header h1 {{
                font-size: 1.4rem;
            }}
            .principle-card {{
                padding: 1.1rem 0.9rem;
            }}
            .card-num, .card-principle-badge {{
                font-size: 0.72rem;
                padding: 0.2rem 0.45rem;
            }}
            .ref-item-chip {{
                font-size: 0.72rem;
                padding: 0.2rem 0.45rem;
            }}
            .flip-card-container {{
                height: 500px;
            }}
            .card-face {{
                padding: 1.1rem 0.85rem;
            }}
        }}
    </style>
</head>
<body>

    <header>
        <div class="controls-container">
            <div>
                <h1>工作的管見</h1>
                <p>職場生存與成長的避坑指南 (中・英・台三語對照卡牌)</p>
            </div>
            
            <div class="action-tools">
                <div class="tabs">
                    <button class="tab-btn active" onclick="switchView('principles')">核心原則 (20)</button>
                    <button class="tab-btn" onclick="switchView('slides')">簡報卡牌 ({len(final_items)})</button>
                    <button class="tab-btn" onclick="switchView('random')">隨機抽卡</button>
                </div>
                
                <div class="search-bar" id="searchBarContainer">
                    <span class="search-icon">🔍</span>
                    <input type="text" class="search-input" id="searchInput" placeholder="搜尋原則、中文、英文或台語..." oninput="handleSearch()">
                </div>
                
                <button class="icon-btn" onclick="toggleDarkMode()" title="切換深淺色模式">🌓</button>
            </div>
        </div>
        
        <div class="display-toggles-bar">
            <div class="toggle-group-left">
                <span class="toggle-group-label">👁️ 顯示切換：</span>
                <div class="toggle-chips">
                    <button class="toggle-chip active" id="toggle-original" onclick="toggleField('original')" title="顯示/關閉投影片原文（目前書寫的原則）">
                        <span class="chip-dot"></span> 簡報原文
                    </button>
                    <button class="toggle-chip active" id="toggle-mandarin" onclick="toggleField('mandarin')" title="顯示/關閉優化中文（國語）">
                        <span class="chip-dot"></span> 優化中文
                    </button>
                    <button class="toggle-chip active" id="toggle-english" onclick="toggleField('english')" title="顯示/關閉優化英文">
                        <span class="chip-dot"></span> 優化英文
                    </button>
                    <button class="toggle-chip active" id="toggle-taiwanese" onclick="toggleField('taiwanese')" title="顯示/關閉優化台文">
                        <span class="chip-dot"></span> 優化台文
                    </button>
                    <button class="toggle-chip active" id="toggle-principles" onclick="toggleField('principles')" title="顯示/關閉核心原則標籤與對照連結">
                        <span class="chip-dot"></span> 對照原則
                    </button>
                </div>
            </div>
            <div class="preset-links">
                <span>快速模式：</span>
                <button class="preset-btn" onclick="applyPreset('all')" title="顯示全部內容">全部顯示</button>
                <button class="preset-btn" onclick="applyPreset('trilingual')" title="只看三語翻譯，隱藏簡報原文">純三語</button>
                <button class="preset-btn" onclick="applyPreset('original_only')" title="只看簡報原文，隱藏三語翻譯">純原文</button>
            </div>
        </div>
    </header>

    <main>
        <!-- VIEW 1: Core Principles (20) -->
        <section id="view-principles" class="section-view active">
            <div class="filter-tags" id="principlesFilters">
                <button class="filter-tag active" onclick="filterPrinciples('all')">全部章節</button>
                <button class="filter-tag" onclick="filterPrinciples('1')">第一篇：效能與工具</button>
                <button class="filter-tag" onclick="filterPrinciples('2')">第二篇：思考與決策</button>
                <button class="filter-tag" onclick="filterPrinciples('3')">第三篇：溝通與協作</button>
                <button class="filter-tag" onclick="filterPrinciples('4')">第四篇：職場生存與防線</button>
                <button class="filter-tag" onclick="filterPrinciples('5')">第五篇：職涯與個人成長</button>
                <button class="filter-tag" onclick="filterPrinciples('6')">第六篇：能量、韌性與生活平衡</button>
            </div>
            <div class="card-grid" id="principlesGrid"></div>
        </section>

        <!-- VIEW 2: Slide Items -->
        <section id="view-slides" class="section-view">
            <div class="filter-tags" id="slidesFilters">
                {slides_filter_html}
            </div>
            <div class="card-grid" id="slidesGrid"></div>
        </section>

        <!-- VIEW 3: Random Card -->
        <section id="view-random" class="section-view">
            <div class="random-view-container">
                <div class="flip-card-container" id="randomCardContainer" onclick="flipRandomCard()">
                    <div class="flip-card">
                        <!-- Front Face -->
                        <div class="card-face card-face-front">
                            <div class="card-header">
                                <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 0.25rem;">
                                    <span class="card-num" id="randomFrontNum">No. 1</span>
                                    <span id="randomFrontPrinciples"></span>
                                </div>
                                <span class="card-meta" id="randomFrontMeta">Slide 1</span>
                            </div>
                            <div class="front-body">
                                <p class="front-title" id="randomFrontOriginal">Click 'Draw Card' to start!</p>
                                <p class="front-placeholder">（簡報原文目前已隱藏，點擊卡牌翻面查看三語對照）</p>
                            </div>
                            <p class="flip-hint">💡 點擊卡牌翻面查看三語對照</p>
                        </div>
                        <!-- Back Face -->
                        <div class="card-face card-face-back">
                            <div class="card-header">
                                <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 0.25rem;">
                                    <span class="card-num" id="randomBackNum">No. 1</span>
                                    <span id="randomBackPrinciples"></span>
                                </div>
                                <span class="card-meta" id="randomBackMeta">三語對照</span>
                            </div>
                            <div class="back-body">
                                <div class="lang-section lang-mandarin" id="randomBackMandarinSection">
                                    <span class="lang-label">優化中文 (Polished Mandarin)</span>
                                    <p class="lang-txt" id="randomBackMandarin">中文翻譯與潤稿內容</p>
                                </div>
                                <div class="lang-section lang-english" id="randomBackEnglishSection">
                                    <span class="lang-label">優化英文 (Polished English)</span>
                                    <p class="lang-txt" id="randomBackEnglish">English polished text</p>
                                </div>
                                <div class="lang-section lang-taiwanese" id="randomBackTaiwaneseSection">
                                    <span class="lang-label">優化台文 (Taiwanese)</span>
                                    <p class="lang-txt" id="randomBackTaiwanese">台語漢字智慧</p>
                                </div>
                                <div class="all-hidden-notice">⚠️ 所有三語翻譯欄位均已關閉，請由上方「顯示切換」開啟</div>
                            </div>
                            <p class="flip-hint">💡 點擊卡牌翻回正面</p>
                        </div>
                    </div>
                </div>
                <button class="draw-btn" onclick="drawRandomCard()">
                    <span>🃏</span> 隨機抽一張卡
                </button>
            </div>
        </section>
    </main>

    <footer>
        <p>© 2026 工作的管見 - 避坑指南與商務工作原則整理與潤稿系統</p>
    </footer>

    <script>
        // Embed parsed JSON data
        const PRINCIPLES = {principles_json};
        const SLIDES = {slides_json};

        let currentActiveView = 'principles';
        let principlesFilter = 'all';
        let slidesFilter = 'all';
        let searchQuery = '';

        const DEFAULT_DISPLAY = {{
            original: true,
            principles: true,
            mandarin: true,
            english: true,
            taiwanese: true
        }};

        let displaySettings = {{ ...DEFAULT_DISPLAY }};

        function loadDisplaySettings() {{
            try {{
                const saved = localStorage.getItem('work_principles_display');
                if (saved) {{
                    displaySettings = Object.assign({{}}, DEFAULT_DISPLAY, JSON.parse(saved));
                }}
            }} catch(e) {{}}
            applyDisplayClasses();
        }}

        function saveDisplaySettings() {{
            try {{
                localStorage.setItem('work_principles_display', JSON.stringify(displaySettings));
            }} catch(e) {{}}
        }}

        function toggleField(field) {{
            displaySettings[field] = !displaySettings[field];
            saveDisplaySettings();
            applyDisplayClasses();
        }}

        function applyPreset(type) {{
            if (type === 'all') {{
                displaySettings = {{ original: true, principles: true, mandarin: true, english: true, taiwanese: true }};
            }} else if (type === 'trilingual') {{
                displaySettings = {{ original: false, principles: true, mandarin: true, english: true, taiwanese: true }};
            }} else if (type === 'original_only') {{
                displaySettings = {{ original: true, principles: true, mandarin: false, english: false, taiwanese: false }};
            }}
            saveDisplaySettings();
            applyDisplayClasses();
        }}

        function applyDisplayClasses() {{
            const fields = ['original', 'principles', 'mandarin', 'english', 'taiwanese'];
            fields.forEach(f => {{
                const btn = document.getElementById(`toggle-${{f}}`);
                if (displaySettings[f]) {{
                    document.body.classList.remove(`hide-${{f}}`);
                    if (btn) btn.classList.add('active');
                }} else {{
                    document.body.classList.add(`hide-${{f}}`);
                    if (btn) btn.classList.remove('active');
                }}
            }});
        }}

        // Initialize UI
        document.addEventListener('DOMContentLoaded', () => {{
            loadDisplaySettings();
            renderPrinciples();
            renderSlides();
            drawRandomCard();
        }});

        function switchView(viewName) {{
            currentActiveView = viewName;
            
            // Toggle active tabs
            const tabButtons = document.querySelectorAll('.tab-btn');
            tabButtons.forEach(btn => btn.classList.remove('active'));
            
            const tabIndexMap = {{ 'principles': 0, 'slides': 1, 'random': 2 }};
            tabButtons[tabIndexMap[viewName]].classList.add('active');
            
            // Toggle view containers
            document.querySelectorAll('.section-view').forEach(view => view.classList.remove('active'));
            document.getElementById(`view-${{viewName}}`).classList.add('active');
            
            // Show/hide search bar based on view
            const searchBar = document.getElementById('searchBarContainer');
            if (viewName === 'random') {{
                searchBar.style.display = 'none';
            }} else {{
                searchBar.style.display = 'flex';
                // Reset search query
                document.getElementById('searchInput').value = '';
                searchQuery = '';
                handleSearch();
            }}
        }}

        function toggleDarkMode() {{
            document.body.classList.toggle('light-mode');
        }}

        function getChapterName(num) {{
            if (num <= 4) return '第一篇：效能與工具';
            if (num <= 8) return '第二篇：思考與決策';
            if (num <= 11) return '第三篇：溝通與協作';
            if (num <= 14) return '第四篇：職場生存與防線';
            if (num <= 17) return '第五篇：職涯與個人成長';
            return '第六篇：能量、韌性與生活平衡';
        }}

        function getChapterId(num) {{
            if (num <= 4) return '1';
            if (num <= 8) return '2';
            if (num <= 11) return '3';
            if (num <= 14) return '4';
            if (num <= 17) return '5';
            return '6';
        }}

        function getMatchedPrinciples(slideKey) {{
            const baseKey = slideKey.split('-')[0];
            const matched = [];
            PRINCIPLES.forEach(p => {{
                const refs = p.ref.split(',').map(r => r.trim());
                if (refs.includes(slideKey) || refs.includes(baseKey)) {{
                    matched.push(p.num);
                }}
            }});
            return matched;
        }}

        function renderPrinciples() {{
            const grid = document.getElementById('principlesGrid');
            grid.innerHTML = '';
            
            const filtered = PRINCIPLES.filter(p => {{
                const chapterId = getChapterId(p.num);
                const matchesChapter = (principlesFilter === 'all' || principlesFilter === chapterId);
                const matchesSearch = !searchQuery || 
                    p.title.toLowerCase().includes(searchQuery) ||
                    p.mandarin.toLowerCase().includes(searchQuery) ||
                    p.english.toLowerCase().includes(searchQuery) ||
                    p.taiwanese.toLowerCase().includes(searchQuery);
                return matchesChapter && matchesSearch;
            }});

            if (filtered.length === 0) {{
                grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; color: var(--text-muted); padding: 3rem;">無符合搜尋條件的原則</div>`;
                return;
            }}

            filtered.forEach(p => {{
                const card = document.createElement('div');
                card.className = 'principle-card';
                card.id = `card-principle-${{p.num}}`;

                const refKeys = p.ref.split(',').map(r => r.trim()).filter(Boolean);
                const refChipsHTML = refKeys.map(key => 
                    `<button class="ref-item-chip" onclick="jumpToSlideCard('${{key}}')" title="點擊直接查看簡報卡牌 No. ${{key}}">No. ${{key}}</button>`
                ).join(' ');

                card.innerHTML = `
                    <div class="card-header">
                        <span class="card-num">原則 ${{p.num}}</span>
                        <span class="card-meta">${{getChapterName(p.num)}}</span>
                    </div>
                    <h3 class="card-title">${{p.title}}</h3>
                    <div class="lang-section lang-mandarin">
                        <span class="lang-label">國語 (Mandarin)</span>
                        <p class="lang-txt">${{p.mandarin}}</p>
                    </div>
                    <div class="lang-section lang-english">
                        <span class="lang-label">English</span>
                        <p class="lang-txt">${{p.english}}</p>
                    </div>
                    <div class="lang-section lang-taiwanese">
                        <span class="lang-label">台語 (Taiwanese)</span>
                        <p class="lang-txt">${{p.taiwanese}}</p>
                    </div>
                    <div class="all-hidden-notice">⚠️ 所有三語翻譯欄位均已關閉</div>
                    <div class="ref-items-container">
                        <span class="ref-items-label">🔗 對照筆記卡牌 (${{refKeys.length}} 條項目，點擊跳轉)</span>
                        <div class="ref-chips-grid">
                            ${{refChipsHTML}}
                        </div>
                    </div>
                `;
                grid.appendChild(card);
            }});
        }}

        function renderSlides() {{
            const grid = document.getElementById('slidesGrid');
            grid.innerHTML = '';

            const filtered = SLIDES.filter(s => {{
                const matchesSlide = (slidesFilter === 'all' || slidesFilter === String(s.slide));
                
                // Get matched principles for this slide
                const matched = getMatchedPrinciples(s.key);
                const matchesPrincipleSearch = matched.some(num => {{
                    const p = PRINCIPLES.find(pr => pr.num === num);
                    if (!p) return false;
                    const searchLower = searchQuery.toLowerCase();
                    return `原則 ${{num}}`.toLowerCase().includes(searchLower) ||
                           `原則${{num}}`.toLowerCase().includes(searchLower) ||
                           p.title.toLowerCase().includes(searchLower);
                }});

                const matchesSearch = !searchQuery || 
                    s.key.toLowerCase().includes(searchQuery) ||
                    s.original.toLowerCase().includes(searchQuery) ||
                    s.mandarin.toLowerCase().includes(searchQuery) ||
                    s.english.toLowerCase().includes(searchQuery) ||
                    s.taiwanese.toLowerCase().includes(searchQuery) ||
                    matchesPrincipleSearch;
                return matchesSlide && matchesSearch;
            }});

            if (filtered.length === 0) {{
                grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; color: var(--text-muted); padding: 3rem;">無符合搜尋條件的卡牌</div>`;
                return;
            }}

            filtered.forEach(s => {{
                const card = document.createElement('div');
                card.className = 'principle-card';
                card.id = `card-slide-${{s.key}}`;
                const matched = getMatchedPrinciples(s.key);
                const badges = matched.map(num => `<span class="card-principle-badge" onclick="jumpToPrinciple(${{num}})" style="cursor:pointer;" title="點擊跳轉至 原則 ${{num}}">原則 ${{num}}</span>`).join(' ');
                card.innerHTML = `
                    <div class="card-header">
                        <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 0.25rem;">
                            <span class="card-num">No. ${{s.key}}</span>
                            ${{badges}}
                        </div>
                        <span class="card-meta">Slide ${{s.slide}}</span>
                    </div>
                    <div class="original-eng">"${{s.original}}"</div>
                    <div class="lang-section lang-mandarin">
                        <span class="lang-label">優化中文 (Polished Mandarin)</span>
                        <p class="lang-txt">${{s.mandarin}}</p>
                    </div>
                    <div class="lang-section lang-english">
                        <span class="lang-label">優化英文 (Polished English)</span>
                        <p class="lang-txt">${{s.english}}</p>
                    </div>
                    <div class="lang-section lang-taiwanese">
                        <span class="lang-label">優化台文 (Taiwanese)</span>
                        <p class="lang-txt">${{s.taiwanese}}</p>
                    </div>
                    <div class="all-hidden-notice">⚠️ 簡報原文與三語翻譯均已關閉</div>
                `;
                grid.appendChild(card);
            }});
        }}

        function jumpToSlideCard(key) {{
            switchView('slides');
            
            slidesFilter = 'all';
            document.querySelectorAll('#slidesFilters .filter-tag').forEach(tag => tag.classList.remove('active'));
            const allTag = document.querySelector('#slidesFilters .filter-tag');
            if (allTag) allTag.classList.add('active');
            
            searchQuery = '';
            const searchInput = document.getElementById('searchInput');
            if (searchInput) searchInput.value = '';
            
            renderSlides();
            
            setTimeout(() => {{
                const targetCard = document.getElementById(`card-slide-${{key}}`);
                if (targetCard) {{
                    targetCard.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                    targetCard.classList.add('highlight-pulse');
                    setTimeout(() => {{
                        targetCard.classList.remove('highlight-pulse');
                    }}, 2500);
                }}
            }}, 120);
        }}

        function jumpToPrinciple(num) {{
            switchView('principles');
            principlesFilter = 'all';
            document.querySelectorAll('#principlesFilters .filter-tag').forEach(tag => tag.classList.remove('active'));
            const allTag = document.querySelector('#principlesFilters .filter-tag');
            if (allTag) allTag.classList.add('active');
            
            searchQuery = '';
            const searchInput = document.getElementById('searchInput');
            if (searchInput) searchInput.value = '';
            
            renderPrinciples();

            setTimeout(() => {{
                const targetCard = document.getElementById(`card-principle-${{num}}`);
                if (targetCard) {{
                    targetCard.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                    targetCard.classList.add('highlight-pulse');
                    setTimeout(() => {{
                        targetCard.classList.remove('highlight-pulse');
                    }}, 2500);
                }}
            }}, 120);
        }}

        function filterPrinciples(chapter) {{
            principlesFilter = chapter;
            document.querySelectorAll('#principlesFilters .filter-tag').forEach(tag => tag.classList.remove('active'));
            event.target.classList.add('active');
            renderPrinciples();
        }}

        function filterSlides(slide) {{
            slidesFilter = slide;
            document.querySelectorAll('#slidesFilters .filter-tag').forEach(tag => tag.classList.remove('active'));
            event.target.classList.add('active');
            renderSlides();
        }}

        function handleSearch() {{
            searchQuery = document.getElementById('searchInput').value.toLowerCase().trim();
            if (currentActiveView === 'principles') {{
                renderPrinciples();
            }} else if (currentActiveView === 'slides') {{
                renderSlides();
            }}
        }}

        // Random Card logic
        function drawRandomCard() {{
            const container = document.getElementById('randomCardContainer');
            container.classList.remove('flipped'); // Reset flip state to show front
            
            // Choose a random slide card
            const randomIndex = Math.floor(Math.random() * SLIDES.length);
            const cardData = SLIDES[randomIndex];
            
            // Create sparkle particles
            createSparkles();

            // Animate card deal
            const innerCard = container.querySelector('.flip-card');
            innerCard.style.transform = 'scale(0.9) rotateY(-10deg)';
            
            setTimeout(() => {{
                // Update text content
                document.getElementById('randomFrontNum').innerText = `No. ${{cardData.key}}`;
                document.getElementById('randomFrontMeta').innerText = `Slide ${{cardData.slide}}`;
                document.getElementById('randomFrontOriginal').innerText = cardData.original;
                
                document.getElementById('randomBackNum').innerText = `No. ${{cardData.key}}`;
                document.getElementById('randomBackMeta').innerText = `Slide ${{cardData.slide}} ・ 對照翻譯`;
                document.getElementById('randomBackMandarin').innerText = cardData.mandarin;
                document.getElementById('randomBackEnglish').innerText = cardData.english;
                document.getElementById('randomBackTaiwanese').innerText = cardData.taiwanese;
                
                // Update matched principles badges
                const matched = getMatchedPrinciples(cardData.key);
                const badges = matched.map(num => `<span class="card-principle-badge">原則 ${{num}}</span>`).join(' ');
                document.getElementById('randomFrontPrinciples').innerHTML = badges;
                document.getElementById('randomBackPrinciples').innerHTML = badges;
                
                innerCard.style.transform = '';
            }}, 200);
        }}

        function flipRandomCard() {{
            const container = document.getElementById('randomCardContainer');
            container.classList.toggle('flipped');
        }}

        function createSparkles() {{
            const container = document.getElementById('view-random');
            const button = document.querySelector('.draw-btn');
            const rect = button.getBoundingClientRect();
            
            // Generate some particle spans
            for (let i = 0; i < 20; i++) {{
                const particle = document.createElement('div');
                particle.className = 'sparkle-particles';
                
                // Color variations
                const colors = ['#6366f1', '#818cf8', '#a5b4fc', '#8b5cf6', '#a78bfa'];
                particle.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
                
                // Position above button
                particle.style.left = `${{rect.left + rect.width/2}}px`;
                particle.style.top = `${{rect.top + window.scrollY}}px`;
                particle.style.opacity = '1';
                particle.style.width = `${{Math.random() * 8 + 4}}px`;
                particle.style.height = particle.style.width;
                
                document.body.appendChild(particle);
                
                // Animate outward
                const angle = Math.random() * Math.PI * 2;
                const distance = Math.random() * 120 + 30;
                const destX = Math.cos(angle) * distance;
                const destY = Math.sin(angle) * distance - 80;
                
                particle.animate([
                    {{ transform: 'translate(0, 0) scale(1)', opacity: 1 }},
                    {{ transform: `translate(${{destX}}px, ${{destY}}px) scale(0)`, opacity: 0 }}
                ], {{
                    duration: Math.random() * 600 + 400,
                    easing: 'cubic-bezier(0.1, 0.8, 0.3, 1)',
                    fill: 'forwards'
                }});
                
                setTimeout(() => {{
                    particle.remove();
                }}, 1000);
            }}
        }}
    </script>
</body>
</html>
"""
    
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print(f"Saved regenerated HTML card web app: index.html")
    
    # Generate ODT handbook as well
    odt_path = "工作原則整理與潤稿.odt"
    generate_odt_handbook(principles, final_items, odt_path)
    
    print("\nWorkflow completed successfully!")

if __name__ == "__main__":
    main()
