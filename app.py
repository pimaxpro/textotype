import os
import subprocess
import tempfile
import streamlit as st

# Cấu hình giao diện trang Streamlit
st.set_page_config(
    page_title="Chuyển đổi LaTeX sang Word",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Chuyển đổi LaTeX sang Word (OMML/MathType)")
st.write(
    "Ứng dụng hỗ trợ chuyển đổi file `.tex` sang file Word `.docx`. "
    "Toàn bộ công thức toán sẽ được tự động chuyển thành định dạng toán học chuẩn của Word (OMML/MathType)."
)

# Khung tải file LaTeX
uploaded_file = st.file_uploader("Tải lên file LaTeX của bạn (.tex)", type=["tex"])

if uploaded_file is not None:
    # Đọc nội dung file
    tex_content = uploaded_file.read().decode("utf-8")
    
    with st.expander("🔍 Xem trước nội dung LaTeX"):
        st.code(tex_content, language="latex")

    # Nút bấm bắt đầu chuyển đổi
    if st.button("🚀 Bắt đầu chuyển đổi", type="primary"):
        with st.spinner("Đang xử lý và chuyển đổi công thức..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, "input.tex")
                output_path = os.path.join(tmpdir, "output.docx")

                # Lưu nội dung vào file tạm
                with open(input_path, "w", encoding="utf-8") as f:
                    f.write(tex_content)

                # Chạy lệnh Pandoc để chuyển đổi
                cmd = ["pandoc", input_path, "-o", output_path]
                result = subprocess.run(cmd, capture_output=True, text=True)

                if result.returncode == 0 and os.path.exists(output_path):
                    st.success("🎉 Chuyển đổi thành công!")
                    
                    # Cho phép người dùng tải file về
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Tải file Word (.docx)",
                            data=f,
                            file_name=uploaded_file.name.replace(".tex", ".docx"),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("❌ Có lỗi xảy ra trong quá trình chuyển đổi:")
                    st.code(result.stderr)
