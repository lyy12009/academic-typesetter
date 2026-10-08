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

st.set_page_config(page_title="學術排版自動化工具鏈", layout="wide")

# 側邊欄導航與快捷連結
st.sidebar.title("🛠️ 學術排版工具選單")
app_mode = st.sidebar.selectbox("請選擇要使用的工具：", [
    "📖 使用說明與 SOP (新手必看)",
    "1. Word 結構萃取器 (Extractor)",
    "2. 橫向排版格式編譯器 (Horizontal)",
    "3. 直向排版格式編譯器 (Vertical)"
])

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎨 排版格式產生器 (GUI)")
st.sidebar.markdown("若需要調整字型、字級或編號格式。")
st.sidebar.markdown("[👉 點擊開啟線上 CSS 排版格式產生器](https://lyy12009.github.io/EssayFormatting/)")


# ==========================================
# 頁面 0：新手必看使用說明與 SOP
# ==========================================
if app_mode == "📖 使用說明與 SOP (新手必看)":
    st.title("🎓 跨維度學術排版自動化系統：操作手冊")
    st.markdown("請嚴格遵循以下四個步驟完成高品質論文與報告！")
    
    st.markdown("---")
    
    col_sop1, col_sop2 = st.columns(2)
    
    with col_sop1:
        st.subheader("📌 步驟一：在 Word 進行極簡結構粗排")
        st.markdown("""
        在 Word 中寫作時，只需專注於文章內容，並遵守以下原則：
        1. **大綱標題：** 直接點擊 Word 上方樣式工具列的 **「標題一」、「標題二」、「標題三」**，絕對不要手動改字體大小來偽裝標題。
        2. **項目與編號：** 使用內建的「項目符號」或「編號」，分層請按 `Tab` 鍵，不要手動打空白鍵對齊。
        3. **文獻引用：** 內文可直接包含夾註（如 `(張春興, 1996)`）。
        4. **圖表與標頭：** 保留圖表標頭文字（如 `圖1 研究架構`），實體物件最後再補。
        5. **暫不處理：** 頁面、註腳、目錄、封面與文末書目此階段先不要理會。
        """)
        
        st.subheader("📌 步驟二：透過 Word 結構萃取器轉出 Markdown")
        st.markdown("""
        1. 點擊左側選單進入 **「1. Word 結構萃取器」**。
        2. 上傳你剛剛寫好的粗排 Word 檔案 (`.docx`)。
        3. 系統會自動將 Word 內建的標題階層與清單轉譯成純淨的 **Markdown 文本**，將其複製備用。
        """)

    with col_sop2:
        st.subheader("📌 步驟三：透過合約與編譯器自動排版")
        st.markdown("""
        1. 前往 [線上 CSS 合約產生器](https://lyy12009.github.io/EssayFormatting/) 挑選範本（如「中原應華碩論」），並複製產生的 **JSON 合約**。
        2. 依文件性質選擇 **「2. 橫排編譯器」** 或 **「3. 直排編譯器」**。
        3. 將 JSON 合約貼入左側輸入框，將 Markdown 文本貼入右側輸入框。
        4. 點擊按鈕執行編譯，即可下載排版完美的 Word 核心檔！
        """)
        
        st.subheader("📌 步驟四：最終定稿與人工補漏")
        st.markdown("""
        1. **補圖表與頁下註：** 在下載的 Word 檔中手動補上圖片、表格與註腳。
        2. **設定頁碼與目錄：** 插入頁碼並利用 Word 的「自動目錄」功能抓取標題。
        3. **書目清單協作：** 將文末原始書目交由 AI (如 Flash Lite 3.5) 排序與校對 APA 格式，貼回後手動套用「懸掛縮排」。
        4. **插入封面**，大功告成！
        """)


# ==========================================
# 模組一：Word 結構萃取器
# ==========================================
elif app_mode == "1. Word 結構萃取器 (Extractor)":
    st.title("📄 Word 結構與清單降維萃取器")
    st.markdown("將學生的 Word 粗排檔（僅點標題與清單）無情剝離視覺雜訊，轉譯為純淨的 Markdown 結構。")
    st.info("💡 **操作提示**：請上傳在 Word 中已點好「標題一、二、三」與清單階層的 `.docx` 檔案。")
    
    uploaded_file = st.file_uploader("請上傳學生的 Word 檔案 (.docx)", type=["docx"])
    
    if uploaded_file is not None:
        doc = docx.Document(uploaded_file)
        markdown_lines = []
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue
            
            style_name = para.style.name
            pPr = para._element.pPr
            
            # 抓取清單 XML 階層
            if pPr is not None and pPr.numPr is not None:
                ilvl_element = pPr.numPr.ilvl
                level = int(ilvl_element.val) if ilvl_element is not None else 0
                indent_spaces = "  " * level
                markdown_lines.append(f"{indent_spaces}- {text}")
                continue
                
            # 抓取標題階層
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
# 共用編譯核心函式
# ==========================================
def parse_hex_color(hex_str):
    hex_str = hex_str.lstrip('#')
    if len(hex_str) == 6: return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))
    return RGBColor(0, 0, 0)

def get_alignment(align_str):
    if align_str == "center": return WD_ALIGN_PARAGRAPH.CENTER
    if align_str == "right": return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT

def get_or_create_style(doc, style_name):
    if style_name in doc.styles: return doc.styles[style_name]
    return doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

def inject_style_contract(doc, json_styles):
    style_mappings = {
        "h1": "Heading 1", "h2": "Heading 2", "h3": "Heading 3",
        "body": "Normal", "olist": "Academic OList", "ulist": "Academic UList"
    }
    for json_key, word_style_name in style_mappings.items():
        if json_key not in json_styles: continue
        contract = json_styles[json_key]
        style_obj = get_or_create_style(doc, word_style_name)
        style_obj.paragraph_format.alignment = get_alignment(contract.get("align", "left"))
        
        # 嚴格防呆：只有內文允許縮排，其餘強制歸零
        if json_key == "body":
            indent_chars = contract.get("indent", 0)
            if indent_chars > 0:
                style_obj.paragraph_format.first_line_indent = Pt(indent_chars * contract.get("size", 13))
            else:
                style_obj.paragraph_format.first_line_indent = Pt(0)
        else:
            style_obj.paragraph_format.first_line_indent = Pt(0)
            
        font = style_obj.font
        font.size = Pt(contract.get("size", 12))
        font.bold = contract.get("bold", False)
        font.italic = contract.get("italic", False)
        font.color.rgb = parse_hex_color(contract.get("color", "#000000"))
        font_name = contract.get("font", "標楷體")
        font.name = font_name
        font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

CHINESE_NUMS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
CHINESE_FORMAL = ["", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾"]

def get_numbering_prefix(level, counts, sys_type):
    c1, c2, c3 = counts
    if sys_type == "none" or not sys_type: return ""
    try:
        if sys_type == "chinese":
            if level == 1: return f"{CHINESE_FORMAL[c1]}、"
            if level == 2: return f"{CHINESE_NUMS[c2]}、"
            if level == 3: return f"（{CHINESE_NUMS[c3]}）"
        elif sys_type == "decimal":
            if level == 1: return f"{c1}. "
            if level == 2: return f"{c1}.{c2} "
            if level == 3: return f"{c1}.{c2}.{c3} "
    except IndexError: return ""
    return ""

def compile_core(md_text, style_contract, num_sys, is_vertical=False):
    doc = Document()
    if is_vertical:
        for section in doc.sections:
            sectPr = section._sectPr
            text_direction = sectPr.find(qn('w:textDirection'))
            if text_direction is None:
                text_direction = OxmlElement('w:textDirection')
                sectPr.append(text_direction)
            text_direction.set(qn('w:val'), 'tbRl')
            
    inject_style_contract(doc, style_contract)
    counts = [0, 0, 0]
    
    for line in md_text.split('\n'):
        original_line = line 
        stripped_line = line.strip()
        if not stripped_line: continue
        
        match_ul = re.match(r'^(\s*)[-*]\s+(.*)', original_line)
        match_ol = re.match(r'^(\s*)(\d+)\.\s+(.*)', original_line)
        
        if stripped_line.startswith('### '):
            counts[2] += 1
            text = stripped_line.replace('### ', '', 1).strip()
            doc.add_paragraph(f"{get_numbering_prefix(3, counts, num_sys)}{text}", style='Heading 3')
        elif stripped_line.startswith('## '):
            counts[1] += 1; counts[2] = 0
            text = stripped_line.replace('## ', '', 1).strip()
            doc.add_paragraph(f"{get_numbering_prefix(2, counts, num_sys)}{text}", style='Heading 2')
        elif stripped_line.startswith('# '):
            counts[0] += 1; counts[1] = 0; counts[2] = 0
            text = stripped_line.replace('# ', '', 1).strip()
            doc.add_paragraph(f"{get_numbering_prefix(1, counts, num_sys)}{text}", style='Heading 1')
        elif match_ul:
            spaces = len(match_ul.group(1))
            list_level = spaces // 2
            text = match_ul.group(2)
            p = doc.add_paragraph(f"• {text}", style='Academic UList')
            p.paragraph_format.left_indent = Pt(24 + (list_level * 18))
        elif match_ol:
            spaces = len(match_ol.group(1))
            list_level = spaces // 2
            num = match_ol.group(2)
            text = match_ol.group(3)
            p = doc.add_paragraph(f"{num}. {text}", style='Academic OList')
            p.paragraph_format.left_indent = Pt(24 + (list_level * 18))
        else:
            doc.add_paragraph(stripped_line, style='Normal')
    return doc


# ==========================================
# 模組二：橫排決定論編譯器
# ==========================================
elif app_mode == "2. 橫排決定論編譯器 (Horizontal)":
    st.title("📚 橫排決定論編譯器 (Horizontal)")
    st.markdown("結合 JSON 樣式合約與 Markdown 文本，瞬間生成標準橫排學術 Word 檔。")
    st.info("💡 尚未產生 JSON 合約嗎？請先至 [線上 CSS 產生器](https://lyy12009.github.io/EssayFormatting/) 挑選範本並複製 JSON。")
    
    col1, col2 = st.columns(2)
    with col1:
        json_input = st.text_area("1. 貼上 JSON 樣式合約：", height=300)
    with col2:
        md_input = st.text_area("2. 貼上 Markdown 文本：", height=300)
        
    if st.button("🚀 執行橫排編譯", type="primary"):
        if not json_input.strip() or not md_input.strip():
            st.error("JSON 合約與 Markdown 文本皆不可為空！")
        else:
            try:
                payload = json.loads(json_input)
                num_sys = payload.get("meta", {}).get("numberingSystem", "chinese")
                style_contract = payload.get("styles", payload)
                doc = compile_core(md_input, style_contract, num_sys, is_vertical=False)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                st.success("✅ 橫排編譯成功！")
                st.download_button("📥 下載橫排 Word 檔案", data=buffer, file_name="Horizontal_Paper.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"編譯失敗，請檢查 JSON 格式是否正確：{e}")


# ==========================================
# 模組三：直排決定論編譯器
# ==========================================
elif app_mode == "3. 直排決定論編譯器 (Vertical)":
    st.title("📜 直排決定論編譯器 (Vertical)")
    st.markdown("專為古典文學、方志與漢語音韻研究設計，自動在 XML 底層注入 `tbRl` 直書流向。")
    st.info("💡 尚未產生 JSON 合約嗎？請先至 [線上 CSS 產生器](https://lyy12009.github.io/EssayFormatting/) 挑選範本並複製 JSON。")
    
    col1, col2 = st.columns(2)
    with col1:
        json_input = st.text_area("1. 貼上 JSON 樣式合約：", height=300)
    with col2:
        md_input = st.text_area("2. 貼上 Markdown 文本：", height=300)
        
    if st.button("🚀 執行直排編譯", type="primary"):
        if not json_input.strip() or not md_input.strip():
            st.error("JSON 合約與 Markdown 文本皆不可為空！")
        else:
            try:
                payload = json.loads(json_input)
                num_sys = payload.get("meta", {}).get("numberingSystem", "chinese")
                style_contract = payload.get("styles", payload)
                doc = compile_core(md_input, style_contract, num_sys, is_vertical=True)
                
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                st.success("✅ 直排編譯成功！")
                st.download_button("📥 下載直排 Word 檔案", data=buffer, file_name="Vertical_Paper.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"編譯失敗，請檢查 JSON 格式是否正確：{e}")
