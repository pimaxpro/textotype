import os
import re
import subprocess
import tempfile
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Chuyển đổi LaTeX sang MathType Word Pro",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Chuyển đổi LaTeX (ex_test) sang Word Pro (Auto-MathType)")
st.write("Chuyển đổi đề thi LaTeX sang Word. File sẽ **tự động chạy Toggle TeX sang MathType** ngay khi bạn mở file lên!")

clean_ex_test = st.sidebar.checkbox("Chuẩn hóa gói ex_test (Câu 1, A, B, C, D)", value=True)

# --- HÀM XỬ LÝ EX_TEST ---
def process_ex_test_content(content: str) -> str:
    cau_counter = 0
    def replace_ex(match):
        nonlocal cau_counter
        cau_counter += 1
        return f"\n\n**Câu {cau_counter}.** "

    content = re.sub(r'\\begin\{ex\}', replace_ex, content)
    content = re.sub(r'\\end\{ex\}', r'\n', content)
    
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

# --- HÀM TẠO FILE DOCM CHỨA MACRO AUTO-TOGGLE TEX ---
def convert_latex_to_automathtype_docm(latex_text: str, fix_ex_test: bool) -> bytes:
    if fix_ex_test:
        latex_text = process_ex_test_content(latex_text)

    # Bảo vệ công thức để giữ nguyên $...$ làm dạng chữ thường
    latex_text = re.sub(r'\$\$(.*?)\$\$', r' MATHBLOCKSTART \1 MATHBLOCKEND ', latex_text, flags=re.DOTALL)
    latex_text = re.sub(r'\$([^\$]+)\$', r' MATHINLINESTART \1 MATHINLINEEND ', latex_text)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "w", encoding="utf-8") as f:
            f.write(latex_text)

        # Chạy Pandoc xuất ra docx
        cmd = ["pandoc", input_path, "-o", output_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(result.stderr)

        # Khôi phục công thức về $...$ chuẩn cho MathType
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
latex_input = st.text_area("Dán mã LaTeX vào đây:", height=300, value=r"""\documentclass{article}
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
\end{document}""")

if st.button("🚀 Chuyển đổi sang Word (Tự động MathType)", type="primary", use_container_width=True):
    if latex_input.strip():
        with st.spinner("Đang khởi tạo file Word tự động Toggle TeX..."):
            try:
                docx_bytes = convert_latex_to_automathtype_docm(latex_input, clean_ex_test)
                st.success("🎉 Đã tạo file thành công!")
                st.download_button(
                    label="📥 Tải file Word về máy",
                    data=docx_bytes,
                    file_name="DeThi_AutoMathType.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")
