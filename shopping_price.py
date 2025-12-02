import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from io import BytesIO
import nltk
from nltk.tokenize import word_tokenize
from nltk import pos_tag

# Ensure NLTK downloads
nltk.download("punkt")
nltk.download("averaged_perceptron_tagger")

# POS groups
POS_GROUPS = {
    "동사": ["VB", "VBD", "VBG", "VBN", "VBP", "VBZ"],
    "명사": ["NN", "NNS", "NNP", "NNPS"],
    "형용사": ["JJ", "JJR", "JJS"],
    "부사": ["RB", "RBR", "RBS"],
}


def should_blank(pos, selected_group):
    if selected_group == "전체":
        return True
    if selected_group in POS_GROUPS:
        return pos in POS_GROUPS[selected_group]
    return False


def generate_test_and_answer(text, pos_group):
    tokens = word_tokenize(text)
    tagged = pos_tag(tokens)

    blank_count = 0
    blanks = {}
    output_words = []

    for word, pos in tagged:
        if should_blank(pos, pos_group) and word.isalpha():
            blank_count += 1
            blanks[blank_count] = word
            output_words.append(f"({blank_count}) ______")
        else:
            output_words.append(word)

    test_text = " ".join(output_words)
    return test_text, blanks


def create_docx(test_text, blanks):
    doc = Document()

    # --- 시험지 헤더 디자인 ---
    table = doc.add_table(rows=2, cols=4)
    table.style = "Table Grid"

    headers = ["반", "이름", "점수", "선생님 확인"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(12)

    for i in range(4):
        table.cell(1, i).text = ""

    doc.add_paragraph("\n")  # spacing

    # 본문 문제
    p = doc.add_paragraph(test_text)
    for run in p.runs:
        run.font.size = Pt(12)

    # --- 정답지 페이지 ---
    doc.add_page_break()
    doc.add_heading("정답지", level=1)

    keys = list(blanks.keys())
    col_len = len(keys) // 3 + 1
    rows = [keys[i:i + col_len] for i in range(0, len(keys), col_len)]

    answers_table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    answers_table.style = "Table Grid"

    for r_idx, row_keys in enumerate(rows):
        for c_idx, k in enumerate(row_keys):
            answers_table.cell(r_idx, c_idx).text = f"{k}. {blanks[k]}"

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------- Streamlit UI ----------------
st.title("📘 연세영어학원 자동 빈칸 출제기")
st.write("업로드한 Word 파일(docx)에서 특정 품사만 골라 자동으로 빈칸 문제 + 정답지를 만들어줍니다.")

uploaded = st.file_uploader("Word 파일 업로드", type=["docx"])
pos_group = st.selectbox("빈칸으로 만들 품사 선택", ["전체", "동사", "명사", "형용사", "부사"])

if uploaded:
    if st.button("시험지 생성하기"):
        doc = Document(uploaded)

        full_text = ""
        for para in doc.paragraphs:
            full_text += para.text + "\n"

        test_text, blanks = generate_test_and_answer(full_text, pos_group)
        output = create_docx(test_text, blanks)

        st.success("시험지 생성 완료!")
        st.download_button(
            "📄 시험지 다운로드",
            data=output,
            file_name="blank_test.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
