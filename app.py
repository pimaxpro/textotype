import os
import re
import subprocess
import tempfile
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Tool Chuẩn Hóa LaTeX sang Word",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Tool Chuẩn Hóa LaTeX (ex_test) sang Word Pro")
st.write("Ứng dụng hỗ trợ chuyển đổi đề thi LaTeX sang Word chuẩn đẹp, xóa bỏ hoàn toàn ký tự lạ.")

# ==========================================
# ⚙️ 1. SIDEBAR CẤU HÌNH CÁC CHỨC NĂNG
# ==========================================
st.sidebar.header("⚙️ Cấu hình xuất file")

math_format = st.sidebar.radio(
    "Định dạng công thức toán:",
    ["Word Equation (OMML)", "MathType (Tự động / Toggle TeX)"],
    help="• Word Equation: Công thức hiển thị sẵn ở dạng toán học chuẩn của Word.\n• MathType: Giữ mã $...$ để chuyển thành MathType khi mở Word."
)

clean_ex_test = st.sidebar.checkbox(
    "Xử lý chuẩn hóa gói ex_test", 
    value=True, 
    help="Tự động đánh số Câu 1, Câu 2... và tách rõ các đáp án A, B, C, D."
)

st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Mẹo cho MathType:** Nếu chọn chế độ MathType, khi mở file Word bạn chỉ cần nhấn **Alt + \\** để toàn bộ công thức tự hóa thành MathType gốc."
)


# ==========================================
# 🛠️ 2. HÀM XỬ LÝ EX_TEST (KHÔNG DÙNG ** MARKDOWN)
# ==========================================
def process_ex_test_content(content: str) -> str:
    """Tự động đếm số câu và định dạng đáp án A, B, C, D (Sử dụng văn bản thuần)"""
    
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
# 🔄 3. HÀM CHUYỂN ĐỔI LATEX SANG WORD & ĐỊNH DẠNG IN ĐẬM
# ==========================================
def convert_latex_to_word(latex_text: str, math_option: str, fix_ex_test: bool) -> bytes:
    if fix_ex_test:
        latex_text = process_ex_test_content(latex_text)
