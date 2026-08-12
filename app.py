import os
import re
import subprocess
import tempfile
import streamlit as st
from docx import Document
from PIL import Image
import google.generativeai as genai

st.set_page_config(
    page_title="Tool Chuẩn Hóa LaTeX & OCR sang Word",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Textotype Pro: OCR Ảnh Toán & Chuẩn Hóa LaTeX sang Word")
st.write("Ứng dụng hỗ trợ OCR ảnh đề thi toán bằng Gemini API và chuyển đổi LaTeX (`ex_test`) sang Word chuẩn đẹp.")

# ==========================================
# 🔑 1. SIDEBAR CẤU HÌNH API & ĐỊNH DẠNG
# ==========================================
st.sidebar.header("🔑 Cấu hình Gemini API")
api_key_input = st.sidebar.text_input(
    "Nhập Gemini API Key của thầy:",
    type="password",
    help="Lấy API Key miễn phí từ Google AI Studio (aistudio.google.com)"
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Cấu hình xuất file Word")

math_format = st.sidebar.radio(
    "Định dạng công thức toán:",
    ["MathType (Tự động / Toggle TeX)", "Word Equation (OMML)"],
    help="• MathType: Giữ mã $...$ để nhấn Alt + \\ biến thành MathType trong Word.\n• Word Equation: Chuyển thành công thức mặc định của Word."
)

clean_ex_test = st.sidebar.checkbox(
    "Xử lý chuẩn hóa gói ex_test", 
    value=True, 
    help="Tự động đánh số Câu 1, Câu 2... và tách rõ các đáp án A, B, C, D."
)

st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Mẹo cho MathType:** Khi mở file Word xuất ra, thầy chỉ cần nhấn **Alt + \\** (hoặc chọn Toggle TeX trong thẻ MathType) để toàn bộ công thức tự hóa thành MathType OLE."
)


# ==========================================
# 🛠️ 2. HÀM XỬ LÝ EX_TEST
# ==========================================
def process_ex_test_content(content: str) -> str:
    """Tự động đếm số câu và định dạng đáp án A, B, C, D"""
    
    cau_counter = 0
    def replace_ex(match):
        nonlocal cau_counter
        cau_counter += 1
        return f"\n\nCâu {cau_counter}. "

    content = re.sub(r'\\begin\{ex\}', replace_ex, content)
    content = re.sub(r'\\end\{ex\}', r'\n', content)
    
    bt_counter = 0
    def replace_bt(match):
        nonlocal bt_counter
        bt_counter += 1
        return f"\n\nBài {bt_counter}. "

    content = re.sub(r'\\begin\{bt\}', replace_bt, content)
    content = re.sub(r'\\end\{bt\}', r'\n', content)

    # Tách 4 đáp án trong lệnh \choice {A}{B}{C}{D}
    choice_pattern = r'\\choice\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
    
    def replace_choice(m):
        a = re.sub(r'\\True\s*', '', m.group(1).strip())
        b = re.sub(r'\\True\s*', '', m.group(2).strip())
        c = re.sub(r'\\True\s*', '', m.group(3).strip())
        d = re.sub(r'\\True\s*', '', m.group(4).strip())
        return f"\n\nA. {a}\n\nB. {b}\n\nC. {c}\n\nD. {d}\n"

    content = re.sub(choice_pattern, replace_choice, content, flags=re.DOTALL)
    content = re.sub(r'\\True\s*', '(Đúng) ', content)

    return content


# ==========================================
# 🧠 3. HÀM OCR ẢNH BẰNG GEMINI API
# ==========================================
def process_image_with_gemini(image: Image.Image, api_key: str) -> str:
    """Gửi ảnh qua Gemini 1.5 Flash để nhận dạng ra mã LaTeX"""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = """
    Bạn là một chuyên gia đánh máy công thức toán học (LaTeX).
    Hãy nhận dạng toàn bộ văn bản và công thức trong bức ảnh sau.
    Yêu cầu:
    1. Giữ nguyên cấu trúc các câu hỏi và các đáp án A, B, C, D.
    2. Nếu cấu trúc câu hỏi có dạng \begin{ex} ... \choice{A}{B}{C}{D} \end{ex} hoặc văn bản thường, hãy giữ nguyên hoặc xuất về định dạng LaTeX chuẩn kẹp công thức trong $...$ (inline) hoặc $$...$$ (display).
    3. Nếu có bảng biến thiên, hãy sử dụng gói tkz-tab để vẽ lại.
    4. Chỉ trả về mã LaTeX, tuyệt đối không giải thích thêm hay bọc mã trong block ```latex.
    """
    
    response = model.generate_content([prompt, image])
    return response.text


# ==========================================
# 🔄 4. HÀM CHUYỂN ĐỔI LATEX SANG WORD
# ==========================================
def convert_latex_to_word(latex_text: str, math_option: str, fix_ex_test: bool) -> bytes:
    if fix_ex_test:
        latex_text = process_ex_test_content(latex_text)

    is_mathtype = "MathType" in math_option

    # Bảo vệ công thức toán nếu chọn chế độ MathType
    if is_mathtype:
        latex_text = re.sub(r'\$\$(.*?)\$\$', r' MATHBLOCKSTART \1 MATHBLOCKEND ', latex_text, flags=re.DOTALL)
        latex_text = re.sub(r'\$([^\$]+)\$', r' MATHINLINESTART \1 MATHINLINEEND ', latex_text)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "w", encoding="utf-8") as f:
            f.write(latex_text)

        # Lệnh Pandoc chuẩn bị thực thi
        cmd = ["pandoc", input_path, "-o", output_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(result.stderr)

        # Đọc file Word xuất ra để xử lý khôi phục công thức và IN ĐẬM
        doc = Document(output_path)
        for p in doc.paragraphs:
            text = p.text
            
            # Khôi phục công thức $...$ cho MathType
            if is_mathtype and ("MATHBLOCKSTART" in text or "MATHINLINESTART" in text):
                text = re.sub(r'MATHBLOCKSTART\s*(.*?)\s*MATHBLOCKEND', r'$$\1$$', text)
                text = re.sub(r'MATHINLINESTART\s*(.*?)\s*MATHINLINEEND', r'$\1$', text)
                p.text = text

            # Tự động tìm và IN ĐẬM các nhãn "Câu X.", "Bài X.", "A.", "B.", "C.", "D."
            pattern = r'^(Câu\s+\d+\.|Bài\s+\d+\.|[A-D]\.)'
            match = re.match(pattern, p.text.strip())
            if match:
                prefix = match.group(1)
                rest = p.text.strip()[len(prefix):]
                p.text = "" 
                
                run_bold = p.add_run(prefix)
                run_bold.bold = True
                p.add_run(rest)

        doc.save(output_path)

        with open(output_path, "rb") as f:
            return f.read()


# ==========================================
# 🖥️ 5. GIAO DIỆN CHÍNH (OCR, EDITOR & UPLOAD)
# ==========================================
tab_ocr, tab1, tab2 = st.tabs([
    "📸 OCR Ảnh Đề Toán (Gemini)", 
    "✍️ Editor Nhập Mã LaTeX", 
    "📁 Tải File LaTeX (.tex)"
])

# Trạng thái mã LaTeX chung giữa các Tab
if "latex_content" not in st.session_state:
    st.session_state.latex_content = r"""\documentclass{article}
\usepackage{ex_test}
\begin{document}

\begin{ex}
    Nghiệm của phương trình $x^2 - 4x + 3 = 0$ là:
    \choice
    {$x = 1; x = 3$}
    {$x = -1; x = -3$}
    {$x = 1; x = -3$}
    {$x = -1; x = 3$}
\end{ex}

\end{document}"""

# --- TAB 1: OCR ẢNH ---
with tab_ocr:
    st.subheader("📸 Tải ảnh hoặc chụp ảnh đề toán:")
    uploaded_img = st.file_uploader("Chọn file ảnh (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])
    
    if uploaded_img is
