import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
import nltk
from nltk.tokenize import word_tokenize
from nltk import pos_tag
import random

# NLTK downloads
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
    return pos in POS_GROUPS[selected_group]


def generate_test_and_answer(text, pos_group, blank_ratio):
    tokens = word_tokenize(text)
    tagged = pos_tag(tokens)

    blank_count = 0
    blanks = {}
    output_words = []

    for word, pos in tagged:
        # 품사 조건 + 비율 조건 + 알파벳 단어만
        if should_blank(pos, pos_group) and word.isalpha():
            if random.random() < blank_ratio:
                blank_count += 1
                blanks[blank_count] = word
                output_words.append(f"({blank_count}) ______")
                continue

        output_words.append(word)

    test_text = " ".join(output_words)
    return test_text, blanks


def create_docx(test_text, blanks):
    doc = Document()

    # ---------------- 시험지 헤더 ----------------
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

    doc.add_paragraph("\n")  # spacing

    # ---------------- 본문 문제 ----------------
    p = doc.add_paragraph(test_text)
    for run in p.runs:
        run.font.size = Pt(12)

    # ---------------- 정답지 ----------------
    doc.add_page_break()
    doc.add_heading("정답지", level=1)

    numbers = list(blanks.keys())
    total = len(numbers)

    # 3열로 나누되, 번호는 "세로 방향"으로 진행하도록
    col_count = 3
    row_count = (total + col_count - 1) // col_count

    # 세로 정렬 구조
    table = doc.add_table(rows=row_count, cols=col_count)
    table.style = "Table Grid"

    index = 1
    for col in range(col_count):
        for row in range(row_count):
            if index <= total:
                key = index
                table.cell(row, col).text = f"{key}. {blanks[key]}"
            index += 1

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------- Streamlit UI ----------------
st.title("📘 연세영어학원 자동 빈칸 출제기")

uploaded = st.file_uploader("Word 파일 업로드", type=["docx"])

pos_group = st.selectbox("빈칸으로 만들 품사 선택", ["전체", "동사", "명사", "형용사", "부사"])

blank_ratio = st.slider("빈칸 생성 비율 (%)", 5, 80, 20)
blank_ratio = blank_ratio / 100

if uploaded:
    if st.button("시험지 생성하기"):
        doc = Document(uploaded)
        full_text = "\n".join([p.text for p in doc.paragraphs])

        test_text, blanks = generate_test_and_answer(full_text, pos_group, blank_ratio)
        output = create_docx(test_text, blanks)

        st.success("시험지 생성 완료!")
        st.download_button(
            "📄 시험지 다운로드",
            data=output,
            file_name="blank_test.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
