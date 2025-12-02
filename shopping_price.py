import streamlit as st
from docx import Document
import random
from io import BytesIO
import nltk

# NLTK 다운로드
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

st.set_page_config(page_title="동사 빈칸 시험지 생성기", layout="wide")
st.title("Verb Blank Test Generator")

st.write("""
업로드한 워드 파일(.docx)의 문장에서 **동사(Verb)**를 랜덤으로 빈칸 처리하여 학습용 시험지를 생성합니다.
""")

# 파일 업로드
uploaded_file = st.file_uploader("워드 파일 업로드 (.docx)", type=["docx"])

# 빈칸 비율 설정
blank_ratio = st.slider("동사 중 빈칸 비율 (%)", min_value=10, max_value=100, value=30, step=5)

# 동사 품사 태그
VERB_TAGS = ['VB', 'VBD', 'VBG', 'VBN', 'VBP', 'VBZ']

def generate_blank_verb_text(text, ratio):
    words = nltk.word_tokenize(text)
    pos_tags = nltk.pos_tag(words)

    # 동사 인덱스 추출
    verb_indices = [i for i, (_, tag) in enumerate(pos_tags) if tag in VERB_TAGS]
    n_blanks = max(1, int(len(verb_indices) * ratio / 100))
    
    if verb_indices:
        blank_indices = random.sample(verb_indices, min(n_blanks, len(verb_indices)))
        for idx in blank_indices:
            words[idx] = "_" * len(words[idx])
    return ' '.join(words)

def process_docx(file, ratio):
    doc = Document(file)
    new_doc = Document()
    
    for para in doc.paragraphs:
        if para.text.strip() != "":
            blank_para = generate_blank_verb_text(para.text, ratio)
            new_doc.add_paragraph(blank_para)
    
    # 파일 저장
    output = BytesIO()
    new_doc.save(output)
    output.seek(0)
    return output

if uploaded_file:
    st.write("✅ 파일 업로드 완료!")
    
    output_file = process_docx(uploaded_file, blank_ratio)
    
    st.download_button(
        label="동사 빈칸 시험지 다운로드",
        data=output_file,
        file_name="verb_blank_test.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
