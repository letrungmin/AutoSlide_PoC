import streamlit as st
import docx
import json
import requests
import time
import os
import io
import urllib.parse
import random
from dotenv import load_dotenv
from openai import OpenAI
from pptx import Presentation

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


# 1. SETTING SYSTEM
st.set_page_config(page_title="OCBS GenAI Slide", page_icon="🪄", layout="wide")

load_dotenv()
LLAMA3_API_KEY = os.getenv("GROQ_API_KEY")
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY")
LLAMA3_BASE_URL = "https://api.groq.com/openai/v1" 

THEME = {
    "bg": RGBColor(255, 255, 255),        
    "title": RGBColor(0, 51, 102),        
    "card_bg": RGBColor(248, 250, 252),   
    "card_border": RGBColor(212, 175, 55),
    "text": RGBColor(51, 65, 85),         
    "accent": RGBColor(0, 153, 117),      
    "chart": [RGBColor(0, 153, 117), RGBColor(212, 175, 55), RGBColor(0, 51, 102), RGBColor(148, 163, 184)]
}
client = OpenAI(api_key=LLAMA3_API_KEY, base_url=LLAMA3_BASE_URL)

# 2. LOGO -> AI GEN -> PEXELS
def download_company_logo(domain, filepath):
    try:
        url = f"https://logo.clearbit.com/{domain}?size=800"
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            with open(filepath, 'wb') as f:
                f.write(res.content)
            return True
    except: pass
    return False

def generate_ai_image(image_prompt, filepath):
    try:
        style = "modern corporate finance, high-tech glass building, glowing digital data charts, professional lighting, cinematic, hyper-realistic, 8k resolution, clean minimalist design --ar 16:9"
        full_prompt = f"{image_prompt}, {style}"
        encoded_prompt = urllib.parse.quote(full_prompt)
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        seed = random.randint(1, 100000)
        url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=1024&height=576&nologo=true&seed={seed}"
        
        res = requests.get(url, headers=headers, timeout=20)
        if res.status_code == 200 and len(res.content) > 5000: 
            with open(filepath, 'wb') as handler:
                handler.write(res.content)
            return True
    except: pass
    return False

def download_image_pexels(keyword, filepath):
    try:
        search_query = f"{keyword} finance business technology"
        res = requests.get(f"https://api.pexels.com/v1/search?query={search_query}&per_page=10&orientation=landscape", headers={"Authorization": PEXELS_API_KEY})
        data = res.json()
        photos = data.get("photos", [])
        if photos:
            photo = random.choice(photos)
            img_data = requests.get(photo["src"]["large"]).content
            with open(filepath, 'wb') as handler:
                handler.write(img_data)
            return True
    except: pass
    return False

def get_image_robust(image_prompt, keyword, filepath):
    if generate_ai_image(image_prompt, filepath): return True
    return download_image_pexels(keyword, filepath)

# 3. AI Core: FORMAT
def get_semantic_chunks_from_docx(file_stream, max_words=600):
    file_stream.seek(0) 
    doc = docx.Document(file_stream)
    chunks, current_chunk, current_word_count = [], [], 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if text: 
            words = text.split()
            current_chunk.append(text + " ") 
            current_word_count += len(words)
            if current_word_count >= max_words:
                chunks.append("\n\n".join(current_chunk))
                current_chunk, current_word_count = [], 0
    if current_chunk: chunks.append("\n\n".join(current_chunk))
    return chunks

def get_slide_json_from_llama3(file_stream, status_container):
    chunks = get_semantic_chunks_from_docx(file_stream)
    all_slides = []
    
    system_prompt = """
    Bạn là Giám đốc Chiến lược tài chính. Chuyển văn bản thành JSON Presentation.
    
    KỶ LUẬT THÉP:
    1. 1 KHỐI DỮ LIỆU = 1 SLIDE DUY NHẤT. Nén thông tin triệt để.
    2. 'bullets': Viết 2-3 ý siêu ngắn.
    3. 'company_domain': Nếu nhắc tới CÔNG TY CỤ THỂ (VD: TCBS, VPBankS...), ghi TÊN MIỀN website (VD: "tcbs.com.vn"). NẾU KHÔNG CÓ, ĐỂ TRỐNG: "".
    4. 'image_prompt': Tự tưởng tượng bối cảnh ảnh minh họa (10 từ tiếng Anh).
    5. 'takeaway': Tự suy luận 1 câu chốt hạ chiến lược sắc bén.
    6. NHÃN BIỂU ĐỒ (chart_data): Cực ngắn (1-3 từ).
    
    PHÂN LOẠI BỐ CỤC:
    - Có năm (2026, 2027) -> "type": "table"
    - Có Tỷ lệ % (45%, 50%) -> "type": "chart"
    - Còn lại -> "type": "text"
    
    JSON MẪU BẮT BUỘC:
    {
        "slides": [
            {
                "title": "Tham Chiếu Thị Trường", 
                "type": "text", 
                "company_domain": "tcbs.com.vn",
                "image_prompt": "A modern stock exchange trading floor",
                "bullets": ["Ý 1", "Ý 2"],
                "takeaway": "Nắm bắt cơ hội để dẫn đầu",
                "speaker_notes": "Nguyên văn"
            }
        ]
    }
    """
    
    progress_bar = status_container.progress(0)
    for i, chunk in enumerate(chunks):
        status_container.write(f"Quá trình đang diễn ra {i+1}/{len(chunks)})...")
        retries = 3
        while retries > 0:
            try:
                res = client.chat.completions.create(
                    model="llama-3.1-8b-instant", 
                    messages=[
                        {"role": "system", "content": system_prompt}, 
                        {"role": "user", "content": f"Khối dữ liệu:\n{chunk}"}
                    ],
                    temperature=0.2, 
                    response_format={"type": "json_object"}
                )
                data = json.loads(res.choices[0].message.content)
                if "slides" in data: all_slides.extend(data["slides"])
                
                progress_bar.progress((i + 1) / len(chunks))
                if i < len(chunks) - 1: time.sleep(3) 
                break
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg or "413" in err_msg:
                    time.sleep(15)
                    retries -= 1
                else: break
    return all_slides if all_slides else None

# 4. ENGINE RENDER
def set_p_format(paragraph, text, size=None, bold=None, color=None, align=None):
    paragraph.text = str(text)
    if align is not None: paragraph.alignment = align
    for run in paragraph.runs:
        if size is not None: run.font.size = size
        if bold is not None: run.font.bold = bold
        if color is not None: run.font.color.rgb = color

def render_pptx_clean(slide_data, template_path, status_container):
    status_container.write("Đang chạy...")
    prs = Presentation(template_path)
    slide_w, slide_h = prs.slide_width, prs.slide_height
    
    SAFE_LEFT = Inches(0.8) 
    SAFE_TOP = Inches(1.2) 
    AVAIL_W = slide_w - Inches(1.6) 
    
    try:
        title_slide = prs.slides.add_slide(prs.slide_layouts[6]) 
        p_title = title_slide.shapes.add_textbox(SAFE_LEFT, Inches(3.0), AVAIL_W, Inches(2.0)).text_frame.paragraphs[0]
        set_p_format(p_title, "CHIẾN LƯỢC CHUYỂN ĐỔI SỐ OCBS", Pt(54), True, THEME["title"], PP_ALIGN.CENTER)
        p_sub = title_slide.shapes.add_textbox(SAFE_LEFT, Inches(4.5), AVAIL_W, Inches(1.0)).text_frame.paragraphs[0]
        set_p_format(p_sub, "Bản Trình Bày", Pt(24), None, THEME["text"], PP_ALIGN.CENTER)
    except: pass

    slides_list = slide_data.get('slides', slide_data) if isinstance(slide_data, dict) else slide_data

    for i, slide_info in enumerate(slides_list):
        if not isinstance(slide_info, dict): continue
        
        try:
            slide_type = str(slide_info.get('type') or 'text')
            title_text = str(slide_info.get('title') or 'Untitled')
            takeaway = str(slide_info.get('takeaway') or '')
            speaker_notes = str(slide_info.get('speaker_notes') or '') 
            
            raw_bullets = slide_info.get('bullets') or []
            bullets = [str(b).strip() for b in raw_bullets if b]
            if not bullets:
                bullets = [speaker_notes[:100] + "..."] if speaker_notes else ["Nội dung đang được cập nhật."]
            if slide_type == 'table':
                t_data = slide_info.get('table_data', [])
                if isinstance(t_data, dict): t_data = [[k, v] for k, v in t_data.items()]
                if not isinstance(t_data, list) or len(t_data) < 2: 
                    slide_type = 'text' 
            elif slide_type == 'chart' and not slide_info.get('chart_data'): 
                slide_type = 'text'

            try: slide = prs.slides.add_slide(prs.slide_layouts[6])
            except: slide = prs.slides.add_slide(prs.slide_layouts[1])
                
            if speaker_notes:
                slide.notes_slide.notes_text_frame.text = f"--- NỘI DUNG NGUYÊN BẢN ---\n{speaker_notes}"

            p_title = slide.shapes.add_textbox(SAFE_LEFT, SAFE_TOP, AVAIL_W, Inches(0.8)).text_frame
            p_title.word_wrap = True
            set_p_format(p_title.paragraphs[0], title_text.upper(), Pt(36), True, THEME["title"])
            
            content_top = Inches(2.2) 
            avail_h = slide_h - content_top - Inches(1.2)
            
            # Layout 1: TEXT
            if slide_type == 'text':
                domain = str(slide_info.get('company_domain', '')).strip()
                img_prompt = str(slide_info.get('image_prompt', 'Modern corporate finance business'))
                img_path = f"temp_gen_img_{int(time.time())}_{i}.png"
                
                has_image = False
                is_logo = False
                
                if domain and len(domain) > 3:
                    if download_company_logo(domain, img_path):
                        has_image, is_logo = True, True
                
                if not has_image:
                    if get_image_robust(img_prompt, "corporate", img_path):
                        has_image = True
                
                layout_style = 'logo_focus' if is_logo else random.choice(['left_text', 'right_text'])
                gap = Inches(0.25) 
                card_w = (AVAIL_W / 2) - Inches(0.2)
                
                # Tọa độ
                if layout_style == 'left_text' or layout_style == 'logo_focus':
                    text_left = SAFE_LEFT
                    img_left = SAFE_LEFT + card_w + Inches(0.4)
                else: # right_text
                    img_left = SAFE_LEFT
                    text_left = SAFE_LEFT + card_w + Inches(0.4)

                #Image/Logo
                if has_image:
                    try:
                        if is_logo:
                            logo_w = Inches(2.5)
                            logo_left = img_left + (card_w - logo_w) / 2
                            slide.shapes.add_picture(img_path, logo_left, content_top + Inches(0.5), width=logo_w)
                        else:
                            slide.shapes.add_picture(img_path, img_left, content_top + Inches(0.2), width=card_w)
                        os.remove(img_path) 
                    except: card_w = AVAIL_W; text_left = SAFE_LEFT 
                else:
                    card_w = AVAIL_W
                    text_left = SAFE_LEFT

                # Text Bento card
                card_h = (avail_h - (gap * (len(bullets) - 1))) / len(bullets)
                for b_idx, bullet_text in enumerate(bullets):
                    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_left, content_top + b_idx * (card_h + gap), card_w, card_h)
                    shape.fill.solid(); shape.fill.fore_color.rgb = THEME["card_bg"]
                    shape.line.color.rgb, shape.line.width = THEME["card_border"], Pt(1.5)
                    
                    tf = shape.text_frame
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.25)
                    tf.word_wrap = True
                    set_p_format(tf.paragraphs[0], bullet_text, Pt(22), None, THEME["text"])

            #TIMELINE
            elif slide_type == 'table':
                t_data = slide_info.get('table_data', [])
                if isinstance(t_data, dict): t_data = [[k, v] for k, v in t_data.items()]
                nodes = t_data[1:] 
                step = AVAIL_W / len(nodes)
                axis_top = Inches(3.5)
                
                line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, SAFE_LEFT, axis_top, AVAIL_W, Inches(0.05))
                line.fill.solid(); line.fill.fore_color.rgb = line.line.color.rgb = THEME["accent"]
                
                for n_idx, node in enumerate(nodes):
                    cx = SAFE_LEFT + (n_idx * step) + (step / 2)
                    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx - Inches(0.12), axis_top - Inches(0.10), Inches(0.24), Inches(0.24))
                    dot.fill.solid(); dot.fill.fore_color.rgb = dot.line.color.rgb = THEME["accent"]
                    
                    p_yr = slide.shapes.add_textbox(cx - Inches(1.0), axis_top - Inches(0.8), Inches(2.0), Inches(0.6)).text_frame.paragraphs[0]
                    set_p_format(p_yr, str(node[0]), Pt(32), True, THEME["title"], PP_ALIGN.CENTER)
                    
                    content_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - ((step - Inches(0.2))/2), axis_top + Inches(0.3), step - Inches(0.2), Inches(2.2))
                    content_box.fill.solid(); content_box.fill.fore_color.rgb = THEME["card_bg"]; content_box.line.color.rgb = THEME["accent"]
                    
                    p_c = content_box.text_frame.paragraphs[0]
                    content_text = str(node[1]) if len(node) > 1 and node[1] is not None else ""
                    set_p_format(p_c, content_text, Pt(18), None, THEME["text"])
                    content_box.text_frame.word_wrap = True

            #CHART
            elif slide_type == 'chart':
                raw_chart_data = slide_info.get('chart_data', {})
                chart_data_dict = {}
                if isinstance(raw_chart_data, dict): chart_data_dict = raw_chart_data
                elif isinstance(raw_chart_data, list):
                    for item in raw_chart_data:
                        if isinstance(item, dict): chart_data_dict.update(item)
                
                if chart_data_dict:
                    clean_dict = {str(k): float(str(v).replace('%', '').replace(',', '').strip()) for k, v in chart_data_dict.items() if str(v).replace('%', '').replace(',', '').strip().replace('.','',1).isdigit()}
                    if clean_dict:
                        max_label, max_val = max(clean_dict.items(), key=lambda x: x[1])
                        display_val = int(max_val) if max_val.is_integer() else max_val
                        
                        kpi_w, chart_w, cmt_w = AVAIL_W * 0.25, AVAIL_W * 0.35, AVAIL_W * 0.35
                        
                        tf_kpi = slide.shapes.add_textbox(SAFE_LEFT, content_top, kpi_w, avail_h).text_frame
                        set_p_format(tf_kpi.paragraphs[0], f"{display_val}%", Pt(80), True, THEME["accent"])
                        set_p_format(tf_kpi.add_paragraph(), f"{str(max_label).upper()}", Pt(22), None, THEME["text"])
                        
                        c_data = CategoryChartData()
                        c_data.categories, _ = list(clean_dict.keys()), c_data.add_series('Tỷ trọng', list(clean_dict.values()))
                        chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, SAFE_LEFT + kpi_w, content_top, chart_w, chart_w, c_data).chart
                        chart.has_legend = False
                        for idx, pt in enumerate(chart.series[0].points): 
                            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = THEME["chart"][idx % len(THEME["chart"])]
                        
                        if bullets:
                            cmt_box = slide.shapes.add_textbox(SAFE_LEFT + kpi_w + chart_w + Inches(0.2), content_top, cmt_w, avail_h).text_frame
                            cmt_box.word_wrap = True
                            set_p_format(cmt_box.paragraphs[0], "\n\n".join(bullets), Pt(22), None, THEME["text"])
                    else:
                        raise ValueError("Dữ liệu Chart không chứa số hợp lệ")

            if takeaway:
                banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, SAFE_LEFT, slide_h - Inches(0.8), AVAIL_W, Inches(0.5))
                banner.fill.solid(); banner.fill.fore_color.rgb = THEME["card_border"]; banner.line.fill.background = None
                p_ban = banner.text_frame.paragraphs[0]
                banner.text_frame.margin_left = Inches(0.2)
                banner.text_frame.word_wrap = True
                display_takeaway = takeaway if "Tự viết" not in takeaway else "Ứng dụng công nghệ để gia tăng lợi thế cạnh tranh cốt lõi."
                set_p_format(p_ban, f"TAKEAWAY: {display_takeaway}", Pt(18), True, THEME["title"])

        except Exception as e:
            err_box = slide.shapes.add_textbox(SAFE_LEFT, Inches(3), AVAIL_W, Inches(2))
            err_box.text_frame.word_wrap = True
            set_p_format(err_box.text_frame.paragraphs[0], f"Trang này bị bỏ qua do AI sinh sai định dạng ({str(e)}). Vui lòng tham khảo Speaker Notes.", Pt(24), True, RGBColor(255,0,0), PP_ALIGN.CENTER)
            status_container.warning(f"Đã bọc lỗi 1 Slide: {str(e)}")

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# 5. STREAMLIT Control
def main():
    st.title("AI AutoSlide Enterprise")
    st.markdown("Hệ thống chuyển đổi Báo cáo sang slide thuyết trình")

    with st.sidebar:
        st.header("Thiết lập Hệ thống")
        template_path = st.text_input("Đường dẫn file Template (.pptx)", "/Users/trungmin/Downloads/AutoSlide_PoC/templates/template_cong_ty_moi.pptx")

    uploaded_file = st.file_uploader("Kéo thả file bản thảo Word (.docx) vào đây", type="docx")

    if uploaded_file is not None:
        if st.button("Kích Hoạt AI Render", use_container_width=True, type="primary"):
            
            with st.status("Hệ thống đang xử lý dữ liệu...", expanded=True) as status:
                try:
                    json_data = get_slide_json_from_llama3(uploaded_file, status)
                    
                    if json_data:
                        pptx_buffer = render_pptx_clean(json_data, template_path, status)
                        status.update(label="Quá trình tạo Slide thành công!", state="complete", expanded=False)
                        
                        st.success("Tác phẩm hoàn thiện! Hãy xem sự xuất hiện của các Logo doanh nghiệp thật và nhịp điệu Trái/Phải của slide mới.")
                        
                        st.download_button(
                            label="Tải Xuống File PowerPoint (.pptx)",
                            data=pptx_buffer,
                            file_name="OCBS_Bao_Cao_Ultimate.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    else:
                        status.update(label="Trích xuất dữ liệu thất bại. AI không thể đọc cấu trúc file.", state="error")
                except Exception as e:
                    status.update(label=f"Xảy ra lỗi hệ thống: {str(e)}", state="error")

if __name__ == "__main__":
    main()