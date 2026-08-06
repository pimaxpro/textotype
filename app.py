import os
import re
import subprocess
import tempfile
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Chuyển đổi LaTeX (ex_test) sang Word Pro",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Chuyển đổi LaTeX (ex_test) sang Word Pro")
st.write("Chuyển đổi đề thi LaTeX (gói `ex_test`) sang Word. Hỗ trợ chuẩn xác Word Equation và MathType (Toggle TeX).")

# --- SIDEBAR CẤU HÌNH ---
st.sidebar.header("⚙️ Cấu hình đầu ra")
math_format = st.sidebar.radio(
    "Định dạng công thức toán:",
    ["Word Equation (OMML)", "MathType (Dạng $...$ để Toggle TeX)"],
    help="• Word Equation: Công thức toán hiển thị sẵn của Microsoft Word.\n• MathType: Giữ nguyên mã $...$ dạng văn bản thô. Khi mở Word, bạn quét chọn và bấm MathType -> Toggle TeX."
)

clean_ex_test = st.sidebar.checkbox(
    "Xử lý chuẩn hóa gói ex_test", 
    value=True, 
    help="Tự động đánh số Câu 1, Câu 2... và gán nhãn A, B, C, D cho các đáp án."
)

# --- HÀM XỬ LÝ GÓI EX_TEST ---
def process_ex_test_content(content: str) -> str:
    """Xử lý cấu trúc ex_test: đếm câu và định dạng đáp án A, B, C, D"""
    cau_counter = 0
    def replace_ex(match):
        nonlocal cau_counter
        cau_counter += 1
        return f"\n\n**Câu {cau_counter}.** "

    content = re.sub(r'\\begin\{ex\}', replace_ex, content)
    content = re.sub(r'\\end\{ex\}', r'\n', content)
    
    bt_counter = 0
    def replace_bt(match):
        nonlocal bt_counter
        bt_counter += 1
        return f"\n\n**Bài {bt_counter}.** "

    content = re.sub(r'\\begin\{bt\}', replace_bt, content)
    content = re.sub(r'\\end\{bt\}', r'\n', content)

    # Xử lý câu lệnh \choice {A}{B}{C}{D}
    choice_pattern = r'\\choice\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
    
    def replace_choice(m):
        a = re.sub(r'\\True\s*', '', m.group(1).strip())
        b = re.sub(r'\\True\s*', '', m.group(2).strip())
        c = re.sub(r'\\True\s*', '', m.group(3).strip())
        d = re.sub(r'\\True\s*', '', m.group(4).strip())
        return f"\n\n**A.** {a}\n\n**B.** {b}\n\n**C.** {c}\n\n**D.** {d}\n"

    content = re.sub(choice_pattern, replace_choice, content, flags=re.DOTALL)
    content = re.sub(r'\\True\s*', '**(Đúng)** ', content)
    return content

# --- HÀM CHUYỂN ĐỔI FILE ---
def convert_latex_to_docx(latex_text: str, math_type_option: str, fix_ex_test: bool) -> bytes:
    # 1. Chuẩn hóa gói ex_test
    if fix_ex_test:
        latex_text = process_ex_test_content(latex_text)

    is_mathtype = "MathType" in math_type_option

    # 2. Nếu chọn MathType: Giấu các công thức $...$ bằng chuỗi an toàn
    if is_mathtype:
        # Thay công thức khối $$...$$ và công thức dòng $...$ bằng chuỗi thay thế
        latex_text = re.sub(r'\$\$(.*?)\$\$', r' MATHBLOCKSTART \1 MATHBLOCKEND ', latex_text, flags=re.DOTALL)
        latex_text = re.sub(r'\$([^\$]+)\$', r' MATHINLINESTART \1 MATHINLINEEND ', latex_text)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "w", encoding="utf-8") as f:
            f.write(latex_text)

        # Chạy Pandoc
        cmd = ["pandoc", input_path, "-o", output_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(result.stderr)

        # 3. Nếu chọn MathType: Đọc lại file Word và đổi chuỗi thay thế thành $...$
        if is_mathtype:
            doc = Document(output_path)
            for p in doc.paragraphs:
                if "MATHBLOCKSTART" in p.text or "MATHINLINESTART" in p.text:
                    new_text = p.text
                    new_text = re.sub(r'MATHBLOCKSTART\s*(.*?)\s*MATHBLOCKEND', r'$$\1$$', new_text)
                    new_text = re.sub(r'MATHINLINESTART\s*(.*?)\s*MATHINLINEEND', r'$\1$', new_text)
                    p.text = new_text
            doc.save(output_path)

        with open(output_path, "rb") as f:
            return f.read()

# --- GIAO DIỆN CHÍNH ---
tab1, tab2 = st.tabs(["✍️ Editor nhập LaTeX", "📁 Upload File (.tex)"])

latex_input = ""

default_code = r"""\documentclass{article}
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

with tab1:
    st.subheader("Dán mã LaTeX vào đây:")
    latex_input = st.text_area("Mã LaTeX", value=default_code, height=320)

with tab2:
    st.subheader("Tải lên file `.tex` của bạn:")
    uploaded_file = st.file_uploader("Chọn file LaTeX", type=["tex"])
    if uploaded_file is not None:
        latex_input = uploaded_file.read().decode("utf-8")
        st.success(f"Đã tải file thành công: {uploaded_file.name}")

# --- NÚT XUẤT FILE ---
st.divider()

if st.button("🚀 Bắt đầu chuyển đổi sang Word", type="primary", use_container_width=True):
    if not latex_input.strip():
        st.warning("⚠️ Vui lòng nhập nội dung LaTeX hoặc tải file lên!")
    else:
        with st.spinner("Đang xử lý nội dung và xuất file Word..."):
            try:
                docx_bytes = convert_latex_to_docx(latex_input, math_format, clean_ex_test)
                
                st.success("🎉 Chuyển đổi thành công!")
                st.download_button(
                    label="📥 Tải file Word (.docx) về máy",
                    data=docx_bytes,
                    file_name="DeThi_Converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error("❌ Đã xảy ra lỗi trong quá trình chuyển đổi:")
                st.code(str(e))
