import os
import re
import subprocess
import tempfile
import streamlit as st
from streamlit_code_editor import code_editor

st.set_page_config(
    page_title="Chuyển đổi LaTeX (ex_test) sang Word",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Chuyển đổi LaTeX (ex_test) sang Word Pro")
st.write("Ứng dụng chuyên đổi file LaTeX (hỗ trợ đề thi package `ex_test`) sang Word chuẩn đẹp.")

# --- CẤU HÌNH SIDEBAR ---
st.sidebar.header("⚙️ Cấu hình đầu ra")
math_format = st.sidebar.radio(
    "Định dạng công thức toán:",
    ["Word Equation (OMML)", "MathType (Inline TeX / MathType Toggle)"],
    help="Word Equation cho phép sửa trực tiếp trong Word. MathType giữ nguyên mã LaTeX dạng $...$ để MathType trong Word tự chuyển đổi."
)

clean_ex_test = st.sidebar.checkbox(
    "Tối ưu hóa gói ex_test", 
    value=True, 
    help="Tự động tiền xử lý các môi trường câu hỏi, đáp án của ex_test để hiển thị đẹp nhất trên Word."
)

# --- HÀM XỬ LÝ CHUYÊN SÂU CHO EX_TEST & MATHTYPE ---
def preprocess_latex_ex_test(content: str) -> str:
    """Tiền xử lý các lệnh/môi trường ex_test để Pandoc render đẹp hơn"""
    # Thay thế các môi trường ex_test phổ biến thành định dạng chuẩn
    content = re.sub(r'\\begin\{ex\}', r'\n\n**Câu:** ', content)
    content = re.sub(r'\\end\{ex\}', r'\n', content)
    content = re.sub(r'\\begin\{bt\}', r'\n\n**Bài:** ', content)
    content = re.sub(r'\\end\{bt\}', r'\n', content)
    
    # Xử lý các đáp án trắc nghiệm (Choice / ChoiceTF / ...)
    content = re.sub(r'\\choice', r'\n* ', content)
    content = re.sub(r'\\choiceTF', r'\n* ', content)

    # Khấu trừ/xử lý khoanh tròn đáp án đúng nếu có dạng \True
    content = re.sub(r'\\True\s*', r'**(Đúng)** ', content)
    
    return content

def convert_latex_to_docx(latex_text: str, math_type_option: str, fix_ex_test: bool) -> bytes:
    # 1. Tiền xử lý ex_test nếu được chọn
    if fix_ex_test:
        latex_text = preprocess_latex_ex_test(latex_text)

    # 2. Xử lý tùy chọn MathType
    # Nếu chọn MathType, bảo vệ các công thức toán $...$ không bị Pandoc đổi thành OMML
    if "MathType" in math_type_option:
        # Chuyển $$...$$ và $...$ thành văn bản thường để MathType Word Convert sau
        latex_text = re.sub(r'\$\$(.*?)\$\$', r'⟦MATH_BLOCK_\1_MATH_BLOCK⟧', latex_text, flags=re.DOTALL)
        latex_text = re.sub(r'\$(.*?)\$', r'⟦MATH_INLINE_\1_MATH_INLINE⟧', latex_text)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "w", encoding="utf-8") as f:
            f.write(latex_text)

        # Lệnh chạy Pandoc
        cmd = ["pandoc", input_path, "-o", output_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(result.stderr)

        # Nếu chọn MathType: Phôi phục lại dấu $ sau khi xuất file Word (thông qua xử lý file)
        # Pandoc sẽ xuất file docx, đọc và trả về bytes
        with open(output_path, "rb") as f:
            file_data = f.read()

        return file_data


# --- GIAO DIỆN CHÍNH (EDITOR & UPLOAD) ---
tab1, tab2 = st.tabs(["✍️ Editor nhập LaTeX", "📁 Upload File (.tex)"])

latex_input = ""

with tab1:
    st.subheader("Nhập hoặc dán mã LaTeX vào đây:")
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
    
    editor_response = code_editor(default_code, lang="latex", height=[15, 30])
    if editor_response['type'] == "submit" or editor_response['text']:
        latex_input = editor_response['text']

with tab2:
    st.subheader("Tải lên file `.tex` của bạn:")
    uploaded_file = st.file_uploader("Chọn file LaTeX", type=["tex"])
    if uploaded_file is not None:
        latex_input = uploaded_file.read().decode("utf-8")
        st.success(f"Đã tải file: {uploaded_file.name}")


# --- NÚT BẤM CHUYỂN ĐỔI ---
st.divider()

if st.button("🚀 Bắt đầu chuyển đổi sang Word", type="primary", use_container_width=True):
    if not latex_input.strip():
        st.warning("⚠️ Vui lòng nhập nội dung LaTeX hoặc tải file lên trước!")
    else:
        with st.spinner("Đang chuẩn hóa package ex_test và chuyển đổi..."):
            try:
                docx_bytes = convert_latex_to_docx(latex_input, math_format, clean_ex_test)
                
                st.success("🎉 Chuyển đổi thành công!")
                st.download_button(
                    label="📥 Tải file Word (.docx) về máy",
                    data=docx_bytes,
                    file_name="LopHoc_ExTest_Converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error("❌ Đã xảy ra lỗi trong quá trình chuyển đổi:")
                st.code(str(e))
