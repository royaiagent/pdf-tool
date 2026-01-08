import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import io

# --- 核心逻辑 ---
def convert_pdf_to_ppt(uploaded_file, conversion_mode, dpi, use_bg_fill):
    # 重置文件指针
    uploaded_file.seek(0)
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    
    # 获取尺寸
    first_page = doc[0]
    width = Pt(first_page.rect.width)
    height = Pt(first_page.rect.height)
    prs.slide_width = width
    prs.slide_height = height

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pages = len(doc)

    for i, page in enumerate(doc):
        progress_bar.progress((i + 1) / total_pages)
        status_text.text(f"正在处理第 {i+1} / {total_pages} 页...")
        
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白页

        # ==========================================
        # 模式 A: 纯图模式 (Visual)
        # ==========================================
        if conversion_mode == "🖼️ 纯图演示模式 (Visual)":
            pix = page.get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("png")
            slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, width=width, height=height)

        # ==========================================
        # 模式 B: 混合模式 (Hybrid) - 背景图 + 文字
        # ==========================================
        elif conversion_mode == "🛡️ 混合编辑模式 (Hybrid)":
            # 1. 先放背景图
            pix = page.get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("png")
            slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, width=width, height=height)
            
            # 2. 再放文字
            extract_text_to_slide(page, slide, use_bg_fill)

        # ==========================================
        # 模式 C: 深度拆解模式 (Deconstructed) - 你的新需求
        # ==========================================
        elif conversion_mode == "🧩 深度拆解模式 (Editable Objects)":
            # 1. 提取并放置所有独立图片 (Images)
            # 获取页面上所有图片的信息
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                # 提取图片字节流
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                
                # 获取图片在页面上的坐标 (Rect)
                # 注意：一张图可能在页面上出现多次，get_image_rects 返回列表
                img_rects = page.get_image_rects(xref)
                
                for rect in img_rects:
                    # 只有当图片有大小时才插入
                    if rect.width > 0 and rect.height > 0:
                        try:
                            slide.shapes.add_picture(
                                io.BytesIO(image_bytes), 
                                Pt(rect.x0), 
                                Pt(rect.y0), 
                                width=Pt(rect.width), 
                                height=Pt(rect.height)
                            )
                        except:
                            pass # 忽略无法处理的极小图片或错误图片

            # 2. 提取并放置文字 (Text)
            # 在拆解模式下，我们强制不加背景色，让文字背景透明
            extract_text_to_slide(page, slide, use_bg_fill=False)

    # 导出
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

def extract_text_to_slide(page, slide, use_bg_fill):
    """提取文字并添加到 PPT 幻灯片的通用函数"""
    text_data = page.get_text("dict")
    for block in text_data["blocks"]:
        if block["type"] == 0: # 文本块
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if not text: continue
                    
                    x0, y0, x1, y1 = span["bbox"]
                    w, h = x1 - x0, y1 - y0
                    
                    # 容错：如果宽高太小，稍微给一点默认值，防止PPT报错
                    if w <= 0: w = 10
                    if h <= 0: h = 10

                    txBox = slide.shapes.add_textbox(Pt(x0), Pt(y0), Pt(w), Pt(h))
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    run = p.add_run()
                    run.text = text
                    run.font.size = Pt(span["size"])
                    
                    # 颜色
                    try:
                        c = span["color"]
                        run.font.color.rgb = RGBColor((c>>16)&0xFF, (c>>8)&0xFF, c&0xFF)
                    except:
                        run.font.color.rgb = RGBColor(0,0,0)

                    # 只有混合模式才需要背景遮挡，拆解模式不需要
                    if use_bg_fill:
                        txBox.fill.solid()
                        txBox.fill.fore_color.rgb = RGBColor(255, 255, 255)

# --- 页面 UI ---
st.set_page_config(page_title="PDF 转 PPT 专业版", layout="wide")
st.title("📄 PDF 转 PPT：专业分层版")

if 'ppt_data' not in st.session_state:
    st.session_state['ppt_data'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = "converted.pptx"

col1, col2 = st.columns([1, 2])

with col1:
    st.info("模式选择")
    mode = st.radio("请选择转换策略：", [
        "🖼️ 纯图演示模式 (Visual)", 
        "🛡️ 混合编辑模式 (Hybrid)", 
        "🧩 深度拆解模式 (Editable Objects)"
    ])
    
    st.markdown("---")
    
    if mode == "🖼️ 纯图演示模式 (Visual)":
        st.caption("也就是“截图转PPT”。100% 还原样子，但里面什么都不能改。")
        dpi = st.slider("清晰度", 100, 300, 150)
        use_bg = False
        
    elif mode == "🛡️ 混合编辑模式 (Hybrid)":
        st.caption("背景是图片，文字覆盖在上面。**样子最还原，且文字可改**，但图片不能移动。")
        dpi = st.slider("背景清晰度", 100, 300, 150)
        use_bg = st.checkbox("文字加白底 (防止重影)", value=True)
        
    elif mode == "🧩 深度拆解模式 (Editable Objects)":
        st.warning("⚠️ 注意：此模式会把图片和文字彻底分开。但复杂的背景装饰（如波浪、渐变色）可能会丢失，变成白底。")
        dpi = 150 # 拆解模式不需要设置背景DPI
        use_bg = False

with col2:
    uploaded_file = st.file_uploader("上传 PDF", type=["pdf"])
    
    if uploaded_file:
        if st.button("🚀 开始转换", type="primary"):
            try:
                with st.spinner("正在逐层拆解 PDF 元素..."):
                    ppt_io = convert_pdf_to_ppt(uploaded_file, mode, dpi, use_bg)
                    st.session_state['ppt_data'] = ppt_io
                    st.session_state['file_name'] = f"{uploaded_file.name.split('.')[0]}_edited.pptx"
                st.success("✅ 处理完成！")
            except Exception as e:
                st.error(f"转换出错: {e}")

    if st.session_state['ppt_data'] is not None:
        st.download_button(
            label="⬇️ 下载最终 PPT",
            data=st.session_state['ppt_data'],
            file_name=st.session_state['file_name'],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
