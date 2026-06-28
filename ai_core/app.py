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
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
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
    
    researcher_prompt = """
    You are a Data Extraction Agent. Analyze the provided text and extract the key facts, statistics, and core logical arguments.
    Organize the extracted information into a structured outline suitable for a professional presentation.
    Do not invent information. Ensure diacritics are preserved.
    """
    
    designer_prompt = """
    You are a Presentation Design Agent. Convert the provided research notes into this EXACT JSON format.
    {
        "slides": [
            {
                "title": "Slide Title",
                "type": "text", 
                "company_domain": "",
                "image_prompt": "10 english words description",
                "takeaway": "Short concluding sentence",
                "bullets": ["Complete idea 1", "Complete idea 2"],
                "table_data": [],
                "chart_data": {}
            }
        ]
    }
    RULES:
    1. TYPE: "table" for timelines, "chart" for proportions, "text" for everything else.
    2. NO EMPTY BULLETS.
    """
    
    for i, chunk in enumerate(chunks):
        status_container.write(f"Agent 1 (Researcher) analyzing data batch {i+1}/{len(chunks)}...")
        try:
            research_res = client.chat.completions.create(
                model="llama-3.1-8b-instant", 
                messages=[{"role": "system", "content": researcher_prompt}, {"role": "user", "content": chunk}],
                temperature=0.1
            )
            research_notes = research_res.choices[0].message.content.strip()
            
            status_container.write(f"Agent 2 (Designer) structuring JSON for batch {i+1}...")
            retries = 3
            while retries > 0:
                try:
                    design_res = client.chat.completions.create(
                        model="llama-3.1-8b-instant", 
                        messages=[{"role": "system", "content": designer_prompt}, {"role": "user", "content": research_notes}],
                        temperature=0.1, response_format={"type": "json_object"}
                    )
                    content = design_res.choices[0].message.content.strip()
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
                        status_container.error(f"Designer Agent error at batch {i+1}: {e}")
        except Exception as e:
            status_container.error(f"Researcher Agent error at batch {i+1}: {e}")
                    
    return {"slides": all_slides}

def render_pptx_clean(slide_data, template_source, report_title, theme):
    prs = Presentation(template_source)
    
    # Render Cover Slide
    try:
        cover_slide = prs.slides.add_slide(prs.slide_layouts[0])
        for shape in cover_slide.placeholders:
            if shape.placeholder_format.type in [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE]:
                shape.text = report_title.upper()
    except Exception:
        pass

    slides_list = slide_data.get('slides', slide_data) if isinstance(slide_data, dict) else slide_data

    for i, s_info in enumerate(slides_list):
        if not isinstance(s_info, dict): continue
        
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
            cleaned_bullets = ["System is summarizing data..."]
        bullets = cleaned_bullets

        # Always use Layout 1 (Title and Content) to inherit Master Slide standards
        layout_idx = 1 if len(prs.slide_layouts) > 1 else 0
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        title_shape = None
        body_shape = None
        
        for shape in slide.placeholders:
            if shape.placeholder_format.type in [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE]:
                title_shape = shape
            elif shape.placeholder_format.type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                body_shape = shape

        if title_shape:
            title_shape.text = title.upper()

        if not body_shape:
            continue

        avail_left = body_shape.left
        avail_top = body_shape.top
        avail_w = body_shape.width
        avail_h = body_shape.height

        if s_type == 'table':
            sp = body_shape.element
            sp.getparent().remove(sp)
            t_data = s_info.get('table_data', [])
            if len(t_data) > 1:
                rows, cols = len(t_data), len(t_data[0])
                table_shape = slide.shapes.add_table(rows, cols, avail_left, avail_top, avail_w, avail_h)
                table = table_shape.table
                for r_idx, row_data in enumerate(t_data):
                    for c_idx, cell_data in enumerate(row_data):
                        table.cell(r_idx, c_idx).text = str(cell_data)

        elif s_type == 'chart':
            sp = body_shape.element
            sp.getparent().remove(sp)
            c_data_dict = s_info.get('chart_data', {})
            try:
                clean = {str(k): float(str(v).replace('%','').replace(',','')) for k, v in c_data_dict.items()}
                c_data = CategoryChartData()
                c_data.categories = list(clean.keys())
                c_data.add_series('Value', list(clean.values()))
                slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, avail_left, avail_top, avail_w, avail_h, c_data)
            except: pass

        if s_type == 'text':
            domain = str(s_info.get('company_domain', '')).strip()
            img_path = f"temp_{i}.png"
            has_img = False
            
            if domain and len(domain) > 3 and download_company_logo(domain, img_path): has_img = True
            elif get_image_robust(str(s_info.get('image_prompt', 'business')), "business", img_path): has_img = True
            
            tf = body_shape.text_frame
            tf.clear() 
            
            if has_img:
                half_w = int(avail_w * 0.5)
                body_shape.width = half_w - Inches(0.2)
                img_left = avail_left + half_w + Inches(0.2)
                try: 
                    slide.shapes.add_picture(img_path, img_left, avail_top, width=half_w - Inches(0.2))
                    os.remove(img_path)
                except: pass

            for b_idx, b_text in enumerate(bullets):
                p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                p.text = b_text
                p.level = 0 

        if takeaway:
            ban_top = prs.slide_height - Inches(0.8)
            ban = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), ban_top, prs.slide_width - Inches(1.0), Inches(0.5))
            ban.fill.solid(); ban.fill.fore_color.rgb = theme["card_border"]; ban.line.fill.background = None
            
            tf = ban.text_frame
            p = tf.paragraphs[0]
            p.text = f"TAKEAWAY: {takeaway}"
            p.font.color.rgb = theme["title"]
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
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
        with st.status("Processing multi-agent pipeline...") as status:
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
                    status.update(label="Pipeline execution complete.", state="complete")
                else: 
                    status.update(label="Agents failed to extract structured data.", state="error")
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