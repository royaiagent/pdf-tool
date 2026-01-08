import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import io

# --- 核心逻辑 (保持不变) ---
def convert_pdf_to_ppt(uploaded_file, conversion_mode, dpi, use_bg_fill):
    # 重置文件指针，确保从头读取
    uploaded_file.seek(0)
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    first_page = doc[0]
    prs.slide_width = Pt(first_page.rect.width)
    prs.slide_height = Pt(first_page.rect.height)

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pages = len(doc)

    for i, page in enumerate(doc):
        progress_bar.progress((i + 1) / total_pages)
        status_text.text(f"正在处理第 {i+1} / {total_pages} 页...")

        # 1. 背景图
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")
        image_stream = io.BytesIO(img_bytes)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # 2. 混合模式文字叠加
        if conversion_mode == "混合编辑模式 (Hybrid)":
            text_data = page.get_text("dict")
            for block in text_data["blocks"]:
                if block["type"] == 0:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span["text"].strip()
                            if not text: continue
                            x0, y0, x1, y1 = span["bbox"]
                            w, h = x1 - x0, y1 - y0
                            
                            txBox = slide.shapes.add_textbox(Pt(x0), Pt(y0), Pt(w), Pt(h))
                            tf = txBox.text_frame
                            tf.word_wrap = True
                            p = tf.paragraphs[0]
                            run = p.add_run()
                            run.text = text
                            run.font.size = Pt(span["size"])
                            
                            try:
                                c = span["color"]
                                run.font.color.rgb = RGBColor((c>>16)&0xFF, (c>>8)&0xFF, c&0xFF)
                            except:
                                run.font.color.rgb = RGBColor(0,0,0)

                            if use_bg_fill:
                                txBox.fill.solid()
                                txBox.fill.fore_color.rgb = RGBColor(255, 255, 255)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- 页面 UI (修复了下载逻辑) ---
st.set_page_config(page_title="PDF 转 PPT 工具", layout="wide")
st.title("📄 超级 PDF 转 PPT 工具")

# 初始化 session state 用于存储转换结果
if 'ppt_data' not in st.session_state:
    st.session_state['ppt_data'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = "converted.pptx"

col1, col2 = st.columns([1, 2])

with col1:
    st.info("设置区域")
    mode = st.radio("选择模式", ["纯图演示模式 (Visual)", "混合编辑模式 (Hybrid)"])
    dpi = st.slider("清晰度", 100, 300, 150)
    use_bg = False
    if mode == "混合编辑模式 (Hybrid)":
        use_bg = st.checkbox("文字加白底 (防重影)", value=True)

with col2:
    uploaded_file = st.file_uploader("请把 PDF 拖进来", type=["pdf"])
    
    if uploaded_file:
        # 只要用户点击了转换，就把结果存到 session_state 里
        if st.button("🚀 开始转换", type="primary"):
            try:
                with st.spinner("正在施展魔法..."):
                    # 执行转换
                    ppt_io = convert_pdf_to_ppt(uploaded_file, mode, dpi, use_bg)
                    
                    # 存入状态
                    st.session_state['ppt_data'] = ppt_io
                    st.session_state['file_name'] = f"{uploaded_file.name.split('.')[0]}_converted.pptx"
                    
                st.success("🎉 转换成功！请点击下方按钮下载")
            except Exception as e:
                st.error(f"出错啦: {e}")

    # 下载按钮单独拿出来，放在 if st.button 外面
    # 这样即使页面刷新，只要 session_state 里有数据，按钮依然有效
    if st.session_state['ppt_data'] is not None:
        st.download_button(
            label="⬇️ 点击下载 PPT",
            data=st.session_state['ppt_data'],
            file_name=st.session_state['file_name'],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
