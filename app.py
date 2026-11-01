import streamlit as st
import docx
import json
import io
import re
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

st.set_page_config(page_title="學術排版半自動化工具鏈", layout="wide")

# 側邊欄導航與快捷連結
st.sidebar.title("🛠️ 學術排版工具選單")
app_mode = st.sidebar.selectbox("請選擇要使用的工具：", [
    "📖 使用說明與 SOP (新手必看)",
    "1. Word 結構萃取器 (Extractor)",
    "2. 橫向排版編譯器 (Horizontal)",
    "3. 直向排版編譯器 (Vertical)"
])

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎨 樣式合約產生器 (GUI)")
st.sidebar.markdown("需要調整字型、字級或編號格式嗎？")
st.sidebar.markdown("[👉 點擊開啟線上 CSS 合約產生器](https://lyy12009.github.io/EssayFormatting/)")


# ==========================================
# 共用編譯核心函式 (底層 OpenXML 雙層字型鎖死)
# ==========================================
def parse_hex_color(hex_str):
    hex_str = str(hex_str).lstrip('#')
    if len(hex_str) == 6:
        return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))
    return RGBColor(0, 0, 0)

def get_alignment(align_str):
    align_str = str(align_str).lower()
    if align_str == "center": return WD_ALIGN_PARAGRAPH.CENTER
    if align_str == "right": return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT

def get_or_create_style(doc, style_name):
    if style_name in doc.styles:
        return doc.styles[style_name]
    return doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

def set_font_force(target_obj, font_name, font_size_pt=None, bold=False, italic=False, color_rgb=None):
    """
    【關鍵核心防禦】
    同時支援 Style 或 Run 物件：
    1. 強制指定 ascii, hAnsi, eastAsia, cs
    2. 加入 w:hint="eastAsia"，強制告知 Word 優先採用東亞字型解碼
    3. 剔除所有 Theme 參照，杜絕 fallback 至 MS Gothic / MS Mincho
    """
    font = target_obj.font
    font.name = font_name
    if font_size_pt is not None:
        font.size = Pt(font_size_pt)
    font.bold = bold
    font.italic = italic
    if color_rgb is not None:
        font.color.rgb = color_rgb

    # 取得或建立底層 XML rPr 節點
    rPr = font._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)

    # 刪除可能存在的干擾主題屬性
    theme_attrs = [
        'w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:csTheme',
        'w:cstheme', 'w:fareast'
    ]
    for attr in theme_attrs:
        if qn(attr) in rFonts.attrib:
            del rFonts.attrib[qn(attr)]

    # 寫入明確字型名稱與解碼 Hint
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:hint'), 'eastAsia')

def inject_style_contract(doc, json_styles):
    style_mappings = {
        "h1": "Heading 1", "h2": "Heading 2", "h3": "Heading 3",
        "h4": "Heading 4", "h5": "Heading 5", "h6": "Heading 6",
        "body": "Normal", "olist": "Academic OList", "ulist": "Academic UList",
        "quote": "Quote"
    }

    body_contract = json_styles.get("body", {"size": 12, "font": "標楷體"})
    default_font = body_contract.get("font", "標楷體")

    for json_key, word_style_name in style_mappings.items():
        contract = json_styles.get(json_key, body_contract if json_key != "quote" else {})
        if not contract and json_key not in ["body", "quote"]:
            continue

        style_obj = get_or_create_style(doc, word_style_name)
        style_obj.paragraph_format.alignment = get_alignment(contract.get("align", "left"))

        if json_key == "body":
            indent_chars = contract.get("indent", 2)
            if indent_chars > 0:
                style_obj.paragraph_format.first_line_indent = Pt(indent_chars * contract.get("size", 12))
            else:
                style_obj.paragraph_format.first_line_indent = Pt(0)
        elif json_key == "quote":
            style_obj.paragraph_format.first_line_indent = Pt(0)
            style_obj.paragraph_format.left_indent = Pt(24)
            style_obj.paragraph_format.right_indent = Pt(24)
        else:
            style_obj.paragraph_format.first_line_indent = Pt(0)

        font_name = contract.get("font", default_font)
        
        # 標題預設字級動態計算
        if json_key.startswith("h"):
            lvl = int(json_key.replace("h", ""))
            default_size = max(12, 18 - (lvl * 1.5))
        else:
            default_size = 12

        font_size = contract.get("size", default_size)
        is_bold = contract.get("bold", True if json_key.startswith("h") else False)
        is_italic = contract.get("italic", False)
        rgb_color = parse_hex_color(contract.get("color", "#000000"))

        # 第一層鎖死：寫入 Style 層
        set_font_force(style_obj, font_name, font_size, bold=is_bold, italic=is_italic, color_rgb=rgb_color)

def add_styled_paragraph(doc, text, style_name, style_contract):
    """
    【第二層鎖死：實體 Run 層級直寫】
    除了套用 Word 樣式外，在段落的第一個 Run 上再次注入東亞字型標籤，
    徹底阻斷 Word 樣式繼承失效或佈景主題 fallback。
    """
    p = doc.add_paragraph(style=style_name)
    if text:
        run = p.add_run(text)
        
        # 尋找對應的 JSON 設定值
        style_key_map = {
            "Heading 1": "h1", "Heading 2": "h2", "Heading 3": "h3",
            "Heading 4": "h4", "Heading 5": "h5", "Heading 6": "h6",
            "Normal": "body", "Quote": "quote",
            "Academic UList": "ulist", "Academic OList": "olist"
        }
        cfg_key = style_key_map.get(style_name, "body")
        contract = style_contract.get(cfg_key, style_contract.get("body", {}))
        
        font_name = contract.get("font", style_contract.get("body", {}).get("font", "標楷體"))
        
        if cfg_key.startswith("h"):
            lvl = int(cfg_key.replace("h", ""))
            default_sz = max(12, 18 - (lvl * 1.5))
        else:
            default_sz = 12
            
        font_size = contract.get("size", default_sz)
        is_bold = contract.get("bold", True if cfg_key.startswith("h") else False)
        is_italic = contract.get("italic", False)
        rgb_color = parse_hex_color(contract.get("color", "#000000"))
        
        set_font_force(run, font_name, font_size, bold=is_bold, italic=is_italic, color_rgb=rgb_color)
    return p

def clean_notebooklm_markdown(text):
    text = text.replace(r'\-', '-')
    text = text.replace(r'\_', '_')
    text = text.replace(r'\*', '*')
    text = text.replace('\\', '')
    return text

def compile_core(md_text, style_contract, num_sys, is_vertical=False):
    doc = docx.Document()
    if is_vertical:
        for section in doc.sections:
            sectPr = section._sectPr
            text_direction = sectPr.find(qn('w:textDirection'))
            if text_direction is None:
                text_direction = OxmlElement('w:textDirection')
                sectPr.append(text_direction)
            text_direction.set(qn('w:val'), 'tbRl')
            
    # 注入樣式定義 (第一層)
    inject_style_contract(doc, style_contract)
    
    for line in md_text.split('\n'):
        original_line = clean_notebooklm_markdown(line)
        stripped_line = original_line.strip()
        if not stripped_line:
            continue
        
        if stripped_line.startswith('|') or stripped_line.startswith('```'):
            continue
        
        match_ul = re.match(r'^(\s*)[-*]\s+(.*)', original_line)
        match_ol = re.match(r'^(\s*)(\d+)\.\s+(.*)', original_line)
        
        # 標題 1 到 6 與引文解析 (雙層鎖死調用)
        if stripped_line.startswith('> '):
            quote_text = stripped_line.replace('> ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, quote_text, 'Quote', style_contract)
        elif stripped_line.startswith('###### '):
            text = stripped_line.replace('###### ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 6', style_contract)
        elif stripped_line.startswith('##### '):
            text = stripped_line.replace('##### ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 5', style_contract)
        elif stripped_line.startswith('#### '):
            text = stripped_line.replace('#### ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 4', style_contract)
        elif stripped_line.startswith('### '):
            text = stripped_line.replace('### ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 3', style_contract)
        elif stripped_line.startswith('## '):
            text = stripped_line.replace('## ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 2', style_contract)
        elif stripped_line.startswith('# '):
            text = stripped_line.replace('# ', '', 1).strip().replace('**', '')
            add_styled_paragraph(doc, text, 'Heading 1', style_contract)
        elif match_ul:
            spaces = len(match_ul.group(1))
            list_level = spaces // 2
            text = match_ul.group(2).replace('**', '')
            p = add_styled_paragraph(doc, f"• {text}", 'Academic UList', style_contract)
            p.paragraph_format.left_indent = Pt(24 + (list_level * 18))
        elif match_ol:
            spaces = len(match_ol.group(1))
            list_level = spaces // 2
            num = match_ol.group(2)
            text = match_ol.group(3).replace('**', '')
            p = add_styled_paragraph(doc, f"{num}. {text}", 'Academic OList', style_contract)
            p.paragraph_format.left_indent = Pt(24 + (list_level * 18))
        else:
            clean_body = stripped_line.replace('**', '')
            add_styled_paragraph(doc, clean_body, 'Normal', style_contract)
            
    return doc


# ==========================================
# 頁面 0：新手必看使用說明與 SOP
# ==========================================
if app_mode == "📖 使用說明與 SOP (新手必看)":
    st.title("🎓 學術排版半自動化系統：操作手冊")
    st.markdown("本系統支援從 Word 萃取或直接貼入 Markdown 文本，自動進行決定論編譯。")
    
    st.markdown("---")
    col_sop1, col_sop2 = st.columns(2)
    with col_sop1:
        st.subheader("📌 步驟一：準備文本")
        st.markdown("""
        1. **Word 粗排：** 使用內建標題（Heading 1-6）與清單。
        2. **NotebookLM 匯出：** 可直接將 AI 生成的 Markdown 貼入系統。
        """)
    with col_sop2:
        st.subheader("📌 步驟二與三：選擇排版方向並編譯")
        st.markdown("""
        1. 貼入 JSON 樣式合約與 Markdown 文本。
        2. 選擇「橫向排版」或「直向排版」一鍵編譯下載。
        """)


# ==========================================
# 模組一：Word 結構萃取器 (Extractor)
# ==========================================
elif app_mode == "1. Word 結構萃取器 (Extractor)":
    st.title("📄 Word 結構與清單萃取器")
    st.markdown("將 Word 初級排版檔案轉譯為純淨的 Markdown 結構。")
    
    uploaded_file = st.file_uploader("請上傳學生的 Word 檔案 (.docx)", type=["docx"])
    
    if uploaded_file is not None:
        doc = docx.Document(uploaded_file)
        markdown_lines = []
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            style_name = para.style.name
            pPr = para._element.pPr
            
            if 'Quote' in style_name or '引文' in style_name:
                markdown_lines.append(f"> {text}")
                continue
                
            if pPr is not None and pPr.numPr is not None:
                ilvl_element = pPr.numPr.ilvl
                level = int(ilvl_element.val) if ilvl_element is not None else 0
                indent_spaces = "  " * level
                markdown_lines.append(f"{indent_spaces}- {text}")
                continue
                
            if style_name.startswith('Heading') or style_name.startswith('標題'):
                try:
                    level_str = re.search(r'\d+', style_name)
                    level = int(level_str.group()) if level_str else 1
                    markdown_lines.append(f"{'#' * level} {text}")
                except ValueError:
                    markdown_lines.append(text)
            else:
                markdown_lines.append(text)
                
        md_output = '\n\n'.join(markdown_lines)
        st.success("✅ 結構萃取成功！")
        st.text_area("請複製以下 Markdown 文本：", value=md_output, height=400)


# ==========================================
# 模組二：橫向排版編譯器 (Horizontal)
# ==========================================
elif app_mode == "2. 橫向排版編譯器 (Horizontal)":
    st.title("📚 橫排編譯器 (Horizontal)")
    st.markdown("結合 JSON 樣式合約與 Markdown 文本，生成標準橫排學術 Word 檔（已套用 Run 層級東亞字型鎖定）。")
    st.info("💡 尚未產生 JSON 樣式格式細節嗎？請先至 [線上 CSS 產生器](https://lyy12009.github.io/EssayFormatting/) 挑選範本並複製 JSON。")
    
    col1, col2 = st.columns(2)
    with col1:
        json_input = st.text_area("1. 貼上 JSON 樣式合約：", height=300)
    with col2:
        md_input = st.text_area("2. 貼上 Markdown 文本：", height=300)
        
    if st.button("🚀 執行橫排編譯", type="primary"):
        if not json_input.strip() or not md_input.strip():
            st.error("JSON 樣式格式細節與 Markdown 文本皆不可為空！")
        else:
            try:
                payload = json.loads(json_input)
                num_sys = payload.get("meta", {}).get("numberingSystem", "chinese")
                style_contract = payload.get("styles", payload)
                doc = compile_core(md_input, style_contract, num_sys, is_vertical=False)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                st.success("✅ 橫排編譯成功！字型已全面鎖定。")
                st.download_button("📥 下載橫排 Word 檔案", data=buffer, file_name="Horizontal_Paper.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"編譯失敗，請檢查 JSON 格式是否正確：{e}")


# ==========================================
# 模組三：直向排版編譯器 (Vertical)
# ==========================================
elif app_mode == "3. 直向排版編譯器 (Vertical)":
    st.title("📜 直向排版編譯器 (Vertical)")
    st.markdown("專為古典文學、方志與漢語音韻研究設計，自動在 XML 底層注入 `tbRl` 直書流向。")
    st.info("💡 尚未產生 JSON 合約嗎？請先至 [線上 CSS 產生器](https://lyy12009.github.io/EssayFormatting/) 挑選範本並複製 JSON。")
    
    col1, col2 = st.columns(2)
    with col1:
        v_json_input = st.text_area("1. 貼上 JSON 樣式合約：", height=300, key="vertical_json_input")
    with col2:
        v_md_input = st.text_area("2. 貼上 Markdown 文本：", height=300, key="vertical_md_input")
        
    if st.button("🚀 執行直排編譯", type="primary", key="vertical_compile_btn"):
        if not v_json_input.strip() or not v_md_input.strip():
            st.error("JSON 合約與 Markdown 文本皆不可為空！")
        else:
            try:
                payload = json.loads(v_json_input)
                num_sys = payload.get("meta", {}).get("numberingSystem", "chinese")
                style_contract = payload.get("styles", payload)
                doc = compile_core(v_md_input, style_contract, num_sys, is_vertical=True)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("✅ 直排編譯器編譯成功！字型已全面鎖定。")
                st.download_button(
                    label="📥 下載直排 Word 檔案", 
                    data=buffer, 
                    file_name="Vertical_Paper.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="vertical_download_btn"
                )
            except Exception as e:
                st.error(f"編譯失敗，請檢查 JSON 格式是否正確：{e}")
