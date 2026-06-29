import streamlit as st
import docx
import json
import requests
import os
import io
import urllib.parse
import random
import PyPDF2
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

st.set_page_config(page_title="Universal AI Presentation Generator", layout="wide")

load_dotenv()

if "GROQ_API_KEY" in st.secrets:
    LLAMA3_API_KEY = st.secrets["GROQ_API_KEY"]
    PEXELS_API_KEY = st.secrets.get("PEXELS_API_KEY", "")
else:
    LLAMA3_API_KEY = os.getenv("GROQ_API_KEY")
    PEXELS_API_KEY = os.getenv("PEXELS_API_KEY")

if not LLAMA3_API_KEY:
    st.error("API Key is missing. Please configure GROQ_API_KEY in Streamlit Secrets.")
    st.stop()

LLAMA3_BASE_URL = "https://api.groq.com/openai/v1" 
client = OpenAI(api_key=LLAMA3_API_KEY, base_url=LLAMA3_BASE_URL)

def hex_to_rgb_color(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))

def download_company_logo(domain, filepath):
    try:
        res = requests.get(f"https://logo.clearbit.com/{domain}?size=800", timeout=10)
        if res.status_code == 200:
            with open(filepath, 'wb') as f: f.write(res.content)
            return True
    except: pass
    return False

def generate_ai_image(image_prompt, filepath):
    try:
        style = "modern corporate finance, high-tech glass building, glowing digital data charts, professional lighting, cinematic, hyper-realistic, 8k resolution, clean minimalist design --ar 16:9"
        encoded = urllib.parse.quote(f"{image_prompt}, {style}")
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(f"https://image.pollinations.ai/prompt/{encoded}?width=1024&height=576&nologo=true&seed={random.randint(1, 100000)}", headers=headers, timeout=20)
        if res.status_code == 200 and len(res.content) > 5000: 
            with open(filepath, 'wb') as f: f.write(res.content)
            return True
    except: pass
    return False

def download_image_pexels(keyword, filepath):
    try:
        res = requests.get(f"https://api.pexels.com/v1/search?query={keyword} finance business&per_page=10&orientation=landscape", headers={"Authorization": PEXELS_API_KEY})
        photos = res.json().get("photos", [])
        if photos:
            img_data = requests.get(random.choice(photos)["src"]["large"]).content
            with open(filepath, 'wb') as f: f.write(img_data)
            return True
    except: pass
    return False

def get_image_robust(prompt, keyword, filepath):
    if generate_ai_image(prompt, filepath): return True
    return download_image_pexels(keyword, filepath)

def chunk_text(text, max_words=600):
    words = text.split()
    chunks = []
    for i in range(0, len(words), max_words):
        chunks.append(" ".join(words[i:i+max_words]))
    return chunks

def get_semantic_chunks_from_docx(file_stream, max_words=600):
    file_stream.seek(0) 
    doc = docx.Document(file_stream)
    text = "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    return chunk_text(text, max_words)

def get_semantic_chunks_from_pdf(file_stream, max_words=600):
    file_stream.seek(0)
    reader = PyPDF2.PdfReader(file_stream)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return chunk_text(text, max_words)

def get_semantic_chunks_from_url(url, max_words=600):
    headers = {"User-Agent": "Mozilla/5.0"}
    res = requests.get(url, headers=headers, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.content, "html.parser")
    
    for script in soup(["script", "style", "nav", "footer", "header"]):
        script.extract()
        
    text = soup.get_text(separator=' ', strip=True)
    return chunk_text(text, max_words)

def get_slide_json_from_llama3(chunks, status_container):
    all_slides = []
    system_prompt = """
    Convert text into JSON Presentation. MUST KEEP DIACRITICS IF ANY.
    MUST RETURN JSON IN THIS EXACT FORMAT:
    {
        "slides": [
            {
                "title": "Slide Title",
                "type": "text", 
                "company_domain": "",
                "image_prompt": "10 english words description",
                "takeaway": "Short concluding sentence",
                "positive_stocks": [],
                "negative_stocks": [],
                "bullets": ["Complete idea 1", "Complete idea 2", "Complete idea 3"],
                "table_data": [],
                "chart_data": {}
            }
        ]
    }
    
    RULES:
    1. TYPE CLASSIFICATION: 
       - Roadmap/Stages/Timeline -> "type": "table", "table_data": [["Time", "Event"]]
       - Proportions/% -> "type": "chart", "chart_data": {"Item 1": 60, "Item 2": 40}
       - Others MUST return -> "type": "text"
    2. NO EMPTY BULLETS & NO INCOMPLETE SENTENCES: 
       - SUMMARIZE LOGICALLY. NEVER LEAVE 'bullets' ARRAY EMPTY [].
       - NEVER CUT OFF SENTENCES. Each element in 'bullets' must be a complete idea.
    """
    
    for i, chunk in enumerate(chunks):
        status_container.write(f"Extracting JSON data (Batch {i+1}/{len(chunks)})...")
        retries = 3
        while retries > 0:
            try:
                res = client.chat.completions.create(
                    model="llama-3.1-8b-instant", 
                    messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": chunk}],
                    temperature=0.2, response_format={"type": "json_object"}
                )
                content = res.choices[0].message.content.strip()
                content = content.replace("```json", "").replace("```", "").strip()
                chunk_json = json.loads(content)
                
                if isinstance(chunk_json, dict):
                    if "slides" in chunk_json: all_slides.extend(chunk_json["slides"])
                    elif "slide" in chunk_json: all_slides.extend(chunk_json["slide"])
                    elif "presentation" in chunk_json: all_slides.extend(chunk_json["presentation"])
                    elif "data" in chunk_json: all_slides.extend(chunk_json["data"])
                    elif "title" in chunk_json: all_slides.append(chunk_json)
                elif isinstance(chunk_json, list): all_slides.extend(chunk_json)
                    
                break
            except Exception as e:
                retries -= 1
                if retries == 0: 
                    status_container.error(f"JSON extraction error at batch {i+1}: {e}")
                    
    return {"slides": all_slides}

def clear_placeholders(slide):
    for shape in slide.placeholders:
        sp = shape.element
        sp.getparent().remove(sp)

def get_dynamic_pt(text, default_size):
    length = len(str(text))
    if length < 30: return Pt(default_size + 4)
    elif length < 80: return Pt(default_size)
    elif length < 150: return Pt(max(12, default_size - 3))
    else: return Pt(max(11, default_size - 6))

def set_p_format(paragraph, text, font_size, bold=False, color_rgb=None, alignment=None):
    paragraph.text = str(text)
    font = paragraph.font
    font.name = 'Arial'
    font.size = font_size
    font.bold = bold
    if color_rgb: font.color.rgb = color_rgb
    if alignment: paragraph.alignment = alignment
    
    if hasattr(paragraph, '_parent'):
        tf = paragraph._parent
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

def add_editable_stock_tag(slide, x, y, stock_code, trend, theme):
    tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(1.1), Inches(0.4))
    bg_color = theme["accent"] if trend == 'up' else RGBColor(220, 38, 38)
    tag.fill.solid(); tag.fill.fore_color.rgb = bg_color; tag.line.fill.background = None
    
    tf = tag.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    set_p_format(tf.paragraphs[0], stock_code.upper(), get_dynamic_pt(stock_code, 12), True, RGBColor(255, 255, 255), PP_ALIGN.CENTER)

def render_pptx_clean(slide_data, template_source, report_title, theme):
    prs = Presentation(template_source)
    SAFE_LEFT, SAFE_TOP = Inches(0.5), Inches(0.8)
    AVAIL_W = prs.slide_width - Inches(1.0)
    
    try:
        ts = prs.slides.add_slide(prs.slide_layouts[6])
        clear_placeholders(ts) 
        set_p_format(ts.shapes.add_textbox(SAFE_LEFT, Inches(3.0), AVAIL_W, Inches(2.0)).text_frame.paragraphs[0], report_title.upper(), Pt(54), True, theme["title"], PP_ALIGN.CENTER)
    except: pass

    slides_list = slide_data.get('slides', slide_data) if isinstance(slide_data, dict) else slide_data

    for i, s_info in enumerate(slides_list):
        if not isinstance(s_info, dict): continue
        try:
            s_type = str(s_info.get('type') or 'text')
            if s_type not in ['table', 'chart']: s_type = 'text'
                
            title = str(s_info.get('title') or 'Content Slide')
            takeaway = str(s_info.get('takeaway') or '')
            
            raw_bullets = [str(b).strip() for b in (s_info.get('bullets') or []) if str(b).strip()]
            cleaned_bullets = []
            for b in raw_bullets:
                if len(b) < 25 and len(cleaned_bullets) > 0:
                    cleaned_bullets[-1] += " " + b
                else:
                    cleaned_bullets.append(b)
                    
            if not cleaned_bullets: 
                cleaned_bullets = ["System is summarizing data...", "Please verify with the original document."]
            bullets = cleaned_bullets
            
            p_stocks = s_info.get('positive_stocks') or []
            n_stocks = s_info.get('negative_stocks') or []

            try: slide = prs.slides.add_slide(prs.slide_layouts[6]); clear_placeholders(slide) 
            except: slide = prs.slides.add_slide(prs.slide_layouts[1]); clear_placeholders(slide)

            p_title = slide.shapes.add_textbox(SAFE_LEFT, SAFE_TOP, AVAIL_W, Inches(0.8)).text_frame
            set_p_format(p_title.paragraphs[0], title.upper(), get_dynamic_pt(title, 32), True, theme["title"])
            
            content_top, avail_h = Inches(1.8), prs.slide_height - Inches(2.8)
            
            curr_x = SAFE_LEFT
            for st_code in p_stocks:
                add_editable_stock_tag(slide, curr_x, content_top, str(st_code), 'up', theme)
                curr_x += Inches(1.2)
            for st_code in n_stocks:
                add_editable_stock_tag(slide, curr_x, content_top, str(st_code), 'down', theme)
                curr_x += Inches(1.2)
                
            if p_stocks or n_stocks: 
                content_top += Inches(0.6) 
                avail_h -= Inches(0.6)

            if s_type == 'table':
                t_data = s_info.get('table_data', [])
                nodes = t_data[1:] if len(t_data) > 1 else t_data
                if not nodes: 
                    s_type = 'text'
                else:
                    step = AVAIL_W / max(len(nodes), 1)
                    axis_y = content_top + Inches(1.0)
                    
                    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, SAFE_LEFT, axis_y, AVAIL_W, Inches(0.05))
                    line.fill.solid(); line.fill.fore_color.rgb = theme["accent"]
                    
                    for n_idx, node in enumerate(nodes):
                        cx = SAFE_LEFT + (n_idx * step) + (step / 2)
                        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx - Inches(0.12), axis_y - Inches(0.1), Inches(0.24), Inches(0.24))
                        dot.fill.solid(); dot.fill.fore_color.rgb = theme["accent"]
                        
                        label_w = step - Inches(0.1)
                        p_yr_tf = slide.shapes.add_textbox(cx - label_w/2, axis_y - Inches(1.2), label_w, Inches(1.0)).text_frame
                        p_yr_tf.word_wrap = True 
                        set_p_format(p_yr_tf.paragraphs[0], str(node[0]), get_dynamic_pt(node[0], 14), True, theme["title"], PP_ALIGN.CENTER)
                        
                        box_w = step - Inches(0.3)
                        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - box_w/2, axis_y + Inches(0.3), box_w, avail_h - Inches(1.4))
                        box.fill.solid(); box.fill.fore_color.rgb = theme["card_bg"]; box.line.color.rgb = theme["accent"]
                        
                        tf = box.text_frame
                        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.1)
                        content_txt = str(node[1]) if len(node)>1 else ""
                        set_p_format(tf.paragraphs[0], content_txt, get_dynamic_pt(content_txt, 14), None, theme["text"])

            elif s_type == 'chart':
                c_data_dict = s_info.get('chart_data', {})
                try:
                    clean = {str(k): float(str(v).replace('%','').replace(',','')) for k, v in c_data_dict.items()}
                    cw, tw = AVAIL_W * 0.45, AVAIL_W * 0.5
                    c_data = CategoryChartData()
                    c_data.categories = list(clean.keys())
                    c_data.add_series('Value', list(clean.values()))
                    slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, SAFE_LEFT, content_top, cw, avail_h, c_data)
                except: s_type = 'text'

            if s_type == 'text':
                card_w = AVAIL_W / 2 - Inches(0.2)
                domain = str(s_info.get('company_domain', '')).strip()
                img_path = f"temp_{i}.png"; has_img = False
                
                if domain and len(domain) > 3 and download_company_logo(domain, img_path): has_img = True
                elif get_image_robust(str(s_info.get('image_prompt', 'business management')), "business", img_path): has_img = True
                
                img_left, text_left = (SAFE_LEFT, SAFE_LEFT + card_w + Inches(0.4)) if i % 2 == 0 else (SAFE_LEFT + card_w + Inches(0.4), SAFE_LEFT)
                if has_img:
                    try: slide.shapes.add_picture(img_path, img_left, content_top, width=card_w); os.remove(img_path)
                    except: card_w, text_left = AVAIL_W, SAFE_LEFT
                else: card_w, text_left = AVAIL_W, SAFE_LEFT

                card_h = avail_h / max(len(bullets), 1)
                for idx, b in enumerate(bullets):
                    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_left, content_top + idx * card_h, card_w, card_h - Inches(0.15))
                    shape.fill.solid(); shape.fill.fore_color.rgb = theme["card_bg"]; shape.line.color.rgb = theme["card_border"]
                    
                    tf = shape.text_frame
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.1)
                    set_p_format(tf.paragraphs[0], b, get_dynamic_pt(b, 16), None, theme["text"])

            if takeaway:
                ban = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, SAFE_LEFT, prs.slide_height - Inches(0.8), AVAIL_W, Inches(0.5))
                ban.fill.solid(); ban.fill.fore_color.rgb = theme["card_border"]; ban.line.fill.background = None
                
                tf = ban.text_frame
                tf.margin_left = tf.margin_right = Inches(0.1)
                set_p_format(tf.paragraphs[0], f"TAKEAWAY: {takeaway}", get_dynamic_pt(takeaway, 16), True, theme["title"], PP_ALIGN.CENTER)

        except Exception as e: continue

    out = io.BytesIO(); prs.save(out); out.seek(0)
    return out

def main():
    st.title("Universal AI Presentation Generator")
    
    if "slide_data" not in st.session_state:
        st.session_state.slide_data = None
    if "edited_data" not in st.session_state:
        st.session_state.edited_data = None
    if "ppt_buffer" not in st.session_state:
        st.session_state.ppt_buffer = None

    with st.sidebar:
        st.header("Report Configuration")
        report_title = st.text_input("Cover Title:", "ANALYSIS REPORT")
        
        template_dir = "ai_core/templates"
        available_templates = {}
        if os.path.exists(template_dir):
            for file in os.listdir(template_dir):
                if file.endswith(".pptx"):
                    display_name = file.replace(".pptx", "").replace("_", " ").title()
                    available_templates[display_name] = os.path.join(template_dir, file)
        
        if not available_templates:
            available_templates["Default Template"] = "ai_core/templates/template_cong_ty_moi.pptx"
            
        selected_template_name = st.selectbox("Select System Template:", list(available_templates.keys()))
        selected_template_path = available_templates[selected_template_name]
        
        custom_template = st.file_uploader("Upload Custom Template (.pptx)", type="pptx")
        
        st.markdown("---")
        st.subheader("Brand Kit")
        col1, col2 = st.columns(2)
        with col1:
            title_hex = st.color_picker("Title", "#003366")
            text_hex = st.color_picker("Text", "#334155")
        with col2:
            accent_hex = st.color_picker("Accent", "#009975")
            border_hex = st.color_picker("Border", "#D4AF37")
            
        custom_theme = {
            "title": hex_to_rgb_color(title_hex),
            "accent": hex_to_rgb_color(accent_hex),
            "card_bg": RGBColor(248, 250, 252), 
            "card_border": hex_to_rgb_color(border_hex),
            "text": hex_to_rgb_color(text_hex),
            "chart": [hex_to_rgb_color(accent_hex), hex_to_rgb_color(border_hex), hex_to_rgb_color(title_hex), RGBColor(148, 163, 184)]
        }
        
    st.subheader("Data Source")
    input_source = st.radio("Select input method:", ["File Upload (DOCX/PDF)", "Website URL"], label_visibility="collapsed")
    
    uf = None
    url_input = ""
    
    if input_source == "File Upload (DOCX/PDF)":
        uf = st.file_uploader("Upload Document", type=["docx", "pdf"])
    else:
        url_input = st.text_input("Enter Web Article URL")
        
    if (uf or url_input) and st.button("1. Analyze Content", type="primary"):
        with st.status("Analyzing source content...") as status:
            try:
                chunks = []
                if uf and uf.name.endswith('.docx'):
                    chunks = get_semantic_chunks_from_docx(uf)
                elif uf and uf.name.endswith('.pdf'):
                    chunks = get_semantic_chunks_from_pdf(uf)
                elif url_input:
                    status.update(label="Fetching website content...")
                    chunks = get_semantic_chunks_from_url(url_input)
                    
                if not chunks:
                    status.update(label="No readable text found in the source.", state="error")
                    st.stop()
                    
                js = get_slide_json_from_llama3(chunks, status)
                if js and len(js.get('slides', [])) > 0:
                    st.session_state.slide_data = js
                    st.session_state.ppt_buffer = None
                    status.update(label="Analysis complete. Please review the slides below.", state="complete")
                else: 
                    status.update(label="System could not extract structured data.", state="error")
            except Exception as e: 
                status.update(label=f"System Error: {str(e)}", state="error")

    if st.session_state.slide_data:
        st.markdown("---")
        st.subheader("2. Review & Edit Content")
        
        edited_slides = []
        for i, slide in enumerate(st.session_state.slide_data['slides']):
            with st.expander(f"Slide {i+1}: {slide.get('title', 'Untitled')}", expanded=False):
                new_title = st.text_input("Title", slide.get('title', ''), key=f"title_{i}")
                new_takeaway = st.text_input("Takeaway", slide.get('takeaway', ''), key=f"takeaway_{i}")
                
                bullets_str = "\n".join(slide.get('bullets', []))
                new_bullets_str = st.text_area("Bullets (one per line)", bullets_str, height=120, key=f"bullets_{i}")
                
                new_slide = slide.copy()
                new_slide['title'] = new_title
                new_slide['takeaway'] = new_takeaway
                new_slide['bullets'] = [b.strip() for b in new_bullets_str.split("\n") if b.strip()]
                edited_slides.append(new_slide)
                
        st.session_state.edited_data = {"slides": edited_slides}

        if st.button("3. Generate PowerPoint"):
            template_source = custom_template if custom_template else selected_template_path
            with st.spinner("Rendering slides..."):
                try:
                    buf = render_pptx_clean(st.session_state.edited_data, template_source, report_title, custom_theme)
                    st.session_state.ppt_buffer = buf
                    st.success("PowerPoint generated successfully.")
                except Exception as e:
                    st.error(f"Render Error: {str(e)}")

    if st.session_state.ppt_buffer:
        st.download_button(
            label="Download PowerPoint File", 
            data=st.session_state.ppt_buffer, 
            file_name="Universal_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )

if __name__ == "__main__":
    main()