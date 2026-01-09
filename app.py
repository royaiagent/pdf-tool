import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn  # 用于注入XML命名空间
import io

# --- 字体强制修正函数 (核心黑科技) ---
def set_font_style(run, font_size, font_color_int):
    """
    强制设置字体为微软雅黑，并保留字号和颜色
    """
    # 1. 设置字号
    run.font.size = Pt(font_size)
    
    # 2. 设置字体名称 (常规设置)
    run.font.name = "Microsoft YaHei"
    
    # 3. 设置中文字体 (底层XML注入，解决PPT不认中文字体的问题)
    # 这一步非常关键，没有它，中文字体往往不会变
    rPr = run.font._element.get_or_add_rPr()
    ea = rPr.get_or_add_ea()
    ea.set(qn('w:eastAsia'), 'Microsoft YaHei')
    
    # 4. 设置颜色
    try:
        # PyMuPDF的颜色有时是整数，有时是列表，做个容错
        if isinstance(font_color_int, int):
            run.font.color.rgb = RGBColor(
                (font_color_int >> 16) & 0xFF, 
                (font_color_int >> 8) & 0xFF, 
                font_color_int & 0xFF
            )
        else:
            run.font.color.rgb = RGBColor(0, 0, 0)
    except:
        run.font.color.rgb = RGBColor(0, 0, 0)

# --- 核心转换逻辑 ---
def convert_pdf_to_ppt(uploaded_file, include_bg_image):
    uploaded_file.seek(0)
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    
    # 获取尺寸
    first_page = doc[0]
    prs.slide_width = Pt(first_page.rect.width)
    prs.slide_height = Pt(first_page.rect.height)

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pages = len(doc)

    for i, page in enumerate(doc):
        progress_bar.progress((i + 1) / total_pages)
        status_text.text(f"正在清洗并重构第 {i+1} / {total_pages} 页文字...")
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # --- 选项：是否保留背景图 ---
        # 如果你只想要纯净的文字版，可以在网页上不勾选这个
        if include_bg_image:
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            # 放入图片作为底层
            slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)

        # --- 核心：文字完美分离与重构 ---
        # 使用 "dict" 模式获取最详细的排版信息
        text_data = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        
        for block in text_data["blocks"]:
            if block["type"] == 0:  # 0 = 文本
                for line in block["lines"]:
                    # 这里我们以“行”为单位创建文本框，保证位置最准
                    # 如果以 block 为单位，段落间距容易乱
                    
                    line_text = ""
                    # 预先计算这一行的边界
                    x0, y0, x1, y1 = line["bbox"]
                    
                    # 创建文本框
                    width = x1 - x0
                    height = y1 - y0
                    if width <= 0: width = 10
                    if height <= 0: height = 10
                    
                    txBox = slide.shapes.add_textbox(Pt(x0), Pt(y0), Pt(width), Pt(height))
                    tf = txBox.text_frame
                    tf.word_wrap = False # 禁止自动换行，因为我们是按行提取的
                    
                    p = tf.paragraphs[0]
                    
                    # 遍历行内的每一个片段(span)
                    for span in line["spans"]:
                        text = span["text"]
                        if not text.strip(): continue
                        
                        run = p.add_run()
                        run.text = text
                        
                        # 调用上面的黑科技函数，强制微软雅黑
                        set_font_style(run, span["size"], span["color"])
                        
                    # 视觉优化：如果是混合模式，给文本框加个半透明白底，避免和背景混在一起看不清
                    # 但你要求“完美分离”，通常意味着背景要是白的。
                    # 这里我做一个智能判断：如果有背景图，就加个白底；如果是纯白背景，就不加。
                    if include_bg_image:
                        txBox.fill.solid()
                        txBox.fill.fore_color.rgb = RGBColor(255, 255, 255)
                        # txBox.fill.transparency = 0.1 # 微微透一点，融合更好（可选）

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- 页面 UI ---
st.set_page_config(page_title="PDF 转 PPT (微软雅黑修正版)", layout="wide")
st.title("📄 PDF 文字完美提取工具")

if 'ppt_data' not in st.session_state:
    st.session_state['ppt_data'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = "converted.pptx"

col1, col2 = st.columns([1, 2])

with col1:
    st.info("设置")
    st.markdown("### 🔠 字体策略")
    st.markdown("已强制启用 **Microsoft YaHei (微软雅黑)** 渲染引擎。所有提取的文字都将规范化为此字体，同时保持原有的字号大小。")
    
    st.markdown("### 🖼️ 背景策略")
    include_bg = st.checkbox("保留原PDF背景图", value=False, help="如果不勾选，PPT背景将是纯白的，只有文字。勾选后，文字会覆盖在图片上（带白色底色）。")

with col2:
    uploaded_file = st.file_uploader("上传 PDF 文件", type=["pdf"])
    
    if uploaded_file:
        if st.button("🚀 开始提取与转换", type="primary"):
            try:
                with st.spinner("正在进行字体规范化处理..."):
                    ppt_io = convert_pdf_to_ppt(uploaded_file, include_bg)
                    st.session_state['ppt_data'] = ppt_io
                    st.session_state['file_name'] = f"{uploaded_file.name.split('.')[0]}_yahei.pptx"
                st.success("✅ 转换完成！文字已转为微软雅黑。")
            except Exception as e:
                st.error(f"发生错误: {e}")

    if st.session_state['ppt_data'] is not None:
        st.download_button(
            label="⬇️ 下载 PPT (微软雅黑版)",
            data=st.session_state['ppt_data'],
            file_name=st.session_state['file_name'],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
