import streamlit as st
from docx import Document
import random
from io import BytesIO
import nltk

nltk.download('punkt')

# --- 페이지 설정 ---
st.set_page_config(page_title="Your Shopping Curator - 빈칸 시험지 생성기", layout="wide")
st.title("📝 Your Blank Test Generator")
st.markdown("""
이 앱은 워드 파일(.docx)에서 텍스트를 불러와 **랜덤으로 단어를 빈칸 처리**하여  
학습용 시험지를 만들어 줍니다.  
- 파일 업로드 후 빈칸 비율을 설정하고 다운로드하세요.
""")

# --- 파일 업로드 ---
uploaded_file = st.file_uploader("📂 워드 파일 업로드 (.docx)", type=["docx"])

# --- 빈칸 비율 ---
blank_ratio = st.slider("빈칸 비율 (%)", min_value=10, max_value=90, value=25, step=5)

# --- 빈칸 생성 함수 ---
def generate_random_blank_text(text, ratio):
    words = nltk.word_tokenize(text)
    n_blanks = max(1, int(len(words) * ratio / 100))
    
    if len(words) > 0:
        blank_indices = random.sample(range(len(words)), min(n_blanks, len(words)))
        for idx in blank_indices:
            words[idx] = "_" * len(words[idx])
    return ' '.join(words)

def process_docx(file, ratio):
    doc = Document(file)
    new_doc = Document()
    
    for para in doc.paragraphs:
        if para.text.strip() != "":
            blank_para = generate_random_blank_text(para.text, ratio)
            new_doc.add_paragraph(blank_para)
    
    # 메모리 상 저장
    output = BytesIO()
    new_doc.save(output)
    output.seek(0)
    return output

# --- 결과 처리 ---
if uploaded_file:
    st.success("✅ 파일 업로드 완료!")
    output_file = process_docx(uploaded_file, blank_ratio)
    
    st.markdown("### 다운로드")
    st.download_button(
        label="⬇️ 빈칸 시험지 다운로드",
        data=output_file,
        file_name="random_blank_test.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# --- 푸터 ---
st.markdown("---")
st.markdown("Made with ❤️ by Your Blank Test Generator")
