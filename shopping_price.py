import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import random
from io import BytesIO
import nltk
from nltk import pos_tag, word_tokenize

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

st.set_page_config(page_title="Your Blank Test Generator", layout="wide")

st.title("📝 Your Blank Test Generator")
st.markdown("업로드한 Word 문서에서 특정 품사만 선택해 랜덤 빈칸 문제와 정답지를 생성합니다.")

uploaded_file = st.file_uploader("📂 Word 파일(.docx) 업로드", type=["docx"])

blank_ratio = st.slider("빈칸 비율 (%)", min_value=5, max_value=80, value=25, step=5)

pos_option = st.selectbox(
    "빈칸으로 만들 품사 선택",
    ["전체", "동사", "명사", "형용사", "부사"]
)

# 품사 매핑
POS_MAP = {
    "동사": ["VB", "VBD", "VBG", "VBN", "VBP", "VBZ"],
    "명사": ["NN", "NNS", "NNP", "NNPS"],
    "형용사": ["JJ", "JJR", "JJS"],
    "부사": ["RB", "RBR", "RBS"]
}

def check_pos(tag, selected):
    if selected == "전체":
        return True
    return tag in POS_MAP[selected]

# 문단 테두리
def set_paragraph_border(p):
    p_pr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement('w:pBdr')

    for border in ['top', 'left', 'bottom', 'right']:
        element = OxmlElement(f'w:{border}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '4')
        element.set(qn('w:color'), '000000')
        p_bdr.append(element)

    p_pr.append(p_bdr)

def process_docx(file, ratio, pos_choice):
    original = Document(file)
    new_doc = Document()

    # 페이지 여백
    section = new_doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # 시험지 상단
    header_table = new_doc.add_table(rows=2, cols=4)
    header_table.style = 'Table Grid'
    header_table.rows[0].cells[0].text = "반:"
    header_table.rows[0].cells[2].text = "이름:"
    header_table.rows[1].cells[0].text = "점수:"
    header_table.rows[1].cells[2].text = "선생님 확인:"
    new_doc.add_paragraph("")

    answer_list = []
    blank_counter = 1  # 번호 (1), (2), (3)...

    for para in original.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        words = word_tokenize(text)
        tagged = pos_tag(words)

        # 선택된 품사의 단어만 후보
        candidates = [i for i, (w, t) in enumerate(tagged) if check_pos(t, pos_choice)]

        if not candidates:
            # 해당 품사 없으면 전체 단어에서 적용
            candidates = list(range(len(words)))

        # 랜덤 빈칸 개수
        n_blanks = max(1, int(len(candidates) * ratio / 100))
        chosen = random.sample(candidates, min(n_blanks, len(candidates)))

        answers = {}

        for idx in chosen:
            original_word = words[idx]
            underline = "_" * len(original_word)
            words[idx] = f"({blank_counter}){underline}"
            answers[blank_counter] = original_word
            blank_counter += 1

        # 새 문단 생성
        new_p = new_doc.add_paragraph(" ".join(words))
        set_paragraph_border(new_p)

        if answers:
            answer_list.append(answers)

    # 정답지 페이지
    new_doc.add_page_break()
    new_doc.add_paragraph("📝 정답지 (Answer Sheet)").bold = True

    for ans_dict in answer_list:
        for num, word in ans_dict.items():
            new_doc.add_paragraph(f"{num}. {word}")

    buffer = BytesIO()
    new_doc.save(buffer)
    buffer.seek(0)
    return buffer


if uploaded_file:
    output = process_docx(uploaded_file, blank_ratio, pos_option)

    st.success("문제가 생성되었습니다.")
    st.download_button(
        "📥 시험지 + 정답지 다운로드",
        data=output,
        file_name="blank_test_with_answers.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
