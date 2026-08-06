import os
import re
import subprocess
import tempfile
import streamlit as st

st.set_page_config(
    page_title="Chuyển đổi LaTeX (ex_test) sang Word Pro",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Chuyển đổi LaTeX (ex_test) sang Word Pro")
st.write("Chuyển đổi đề thi LaTeX (gói `ex_test`) sang Word chuẩn đẹp, hỗ trợ cả Word Equation và MathType.")

# --- SIDEBAR CẤU HÌNH ---
st.sidebar.header("⚙️ Cấu hình đầu ra")
math_format = st.sidebar.radio(
    "Định dạng công thức toán:",
    ["Word Equation (OMML - Khuyên dùng)", "MathType (Giữ $...$ để Toggle TeX)"],
    help="• Word Equation: Công thức toán hiển thị sẵn trong Word (Sửa trực tiếp hoặc dùng MathType -> Convert Equations).\n• MathType: Giữ mã $...$ dưới dạng chữ thường để quét Toggle TeX trong Word."
)

clean_ex_test = st.sidebar.checkbox(
    "Xử lý chuẩn hóa gói ex_test", 
    value=True, 
    help="Tự động đánh số Câu 1, Câu 2... và gán nhãn A, B, C, D cho các đáp án."
)


# --- HÀM XỬ LÝ GÓI EX_TEST ---
def process_ex_test_content(content: str) -> str:
    """Xử lý cấu trúc ex_test: đếm câu và định dạng đáp án A, B, C, D"""
    
    # 1. Đếm và thay thế môi trường \begin{ex}...\end{ex} hoặc \begin{bt}...\end{bt}
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

    # 2. Xử lý câu lệnh \choice {A}{B}{C}{D}
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


# --- HÀM CHUYỂN ĐỔI SANG DOCX ---
def convert_latex_to_docx(latex_text: str, math_type_option: str, fix_ex_test: bool) -> bytes:
    # 1. Xử lý ex_test nếu chọn
    if fix_ex_test:
        latex_text = process_ex_test_content(latex_text)

    # 2. Xử lý chế độ xuất file
    is_mathtype = "MathType" in math_type_option

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        if is_mathtype:
            # Đối với MathType: Biến đổi file .tex thành file Markdown để Pandoc không tự động ép công thức thành OMML
            # Đổi $...$ thành dạng văn bản thuần
            latex_text = latex_text.replace(r'\
