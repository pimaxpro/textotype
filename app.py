import os
import re
import subprocess
import tempfile
import streamlit as st

st.set_page_config(
    page_title="Chuyển đổi LaTeX (ex_test) sang Word",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Chuyển đổi LaTeX (ex_test) sang Word Pro")
st.write("Ứng dụng chuyển đổi file LaTeX (hỗ trợ đề thi package `ex_test`) sang Word chuẩn đẹp.")

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

# --- HÀM XỬ LÝ EX_TEST & MATHTYPE ---
def preprocess_latex_ex_test(content: str) -> str:
    """Tiền xử lý các lệnh/môi trường ex_test để Pandoc render đẹp hơn"""
    content = re.sub(r'\\begin\{ex\}', r'\n\n**Câu:** ', content)
    content = re.sub(r'\\end\{ex\}', r'\n', content)
    content = re.sub(r'\\begin\{bt\}', r'\n\n**Bài:** ', content)
    content = re.sub(r'\\end\{bt\}', r'\n', content)
    content = re.sub(r'\\choice', r'\n* ', content)
    content = re.sub(r'\\choiceTF', r'\n* ', content)
    content = re.sub(r'\\True\s*', r'**(Đúng)** ', content)
    return content

def convert_latex_to_docx(latex_text: str, math_type_option: str, fix_ex_test: bool) -> bytes:
    if fix_ex_test:
        latex_text = preprocess_latex_ex_test(latex_text)

    if "MathType" in math_type_option:
        latex_text = re.sub(r'\$\$(.*?)\$\$', r'⟦MATH_BLOCK_\1_MATH_BLOCK⟧', latex_text, flags=re.DOTALL)
        latex_text = re.sub(r'\$(.*?)\$', r'⟦MATH_INLINE_\1_MATH_INLINE⟧', latex_text)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.tex")
        output_path = os.path.join(tmpdir, "output.docx")

        with open(input_path, "w", encoding="utf-8") as f:
            f.write(latex_text)

        cmd = ["pandoc", input_path, "-o", output_path]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise Exception(result.stderr)

        with open(output_path, "rb") as f:
            file_data = f.read()

        return file_data


# --- GIAO DIỆN CHÍNH (TAB EDITOR & TAB UPLOAD) ---
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
    latex_input = st.text_area("Mã LaTeX", value=default_code, height=300)

with tab2:
    st.subheader("Tải lên file `.tex` của bạn:")
    uploaded_file = st.file_uploader("Chọn file LaTeX", type=["tex"])
    if uploaded_file is not None:
        latex_input = uploaded_file.read().decode("utf-8")
        st.success(f"Đã tải file thành công: {uploaded_file.name}")


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
                    file_name="ExTest_Converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error("❌ Đã xảy ra lỗi trong quá trình chuyển đổi:")
                st.code(str(e))
