import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from io import BytesIO
import nltk
from nltk import pos_tag, word_tokenize
import random
import re
import math
import os

# ---------- NLTK data ----------
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

try:
    nltk.data.find('taggers/averaged_perceptron_tagger')
except LookupError:
    nltk.download('averaged_perceptron_tagger', quiet=True)

# ---------- POS 그룹 ----------
POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

TOKEN_CANDIDATE_RE = re.compile(r"[A-Za-z0-9\uac00-\ud7a3]+")

def is_candidate_token(tok):
    return bool(TOKEN_CANDIDATE_RE.search(tok))

def tokenize_preserve_spacing(text):
    tokens = word_tokenize(text)
    return tokens

def assemble_tokens(tokens):
    out = ""
    for i, t in enumerate(tokens):
        if i == 0:
            out += t
            continue
        if re.fullmatch(r"[^\w\s]", t):
            out += t
        else:
            out += " " + t
    return out

def set_runs_font(paragraph, size_pt=11, bold=False):
    for run in paragraph.runs:
        run.font.size = Pt(size_pt)
        run.font.bold = bold

def process_docx_with_answer(file_like, pos_choice, blank_ratio_fraction):
    src = Document(file_like)
    dst = Document()

    # 여백
    for section in dst.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # ---------- 상단 학원 이름 ----------
    title_p = dst.add_paragraph("연세영어학원")
    set_runs_font(title_p, size_pt=18, bold=True)
    title_p.alignment = 1  # 가운데
    dst.add_paragraph("")

    # ---------- 깔끔한 정보란 ----------
    info_text = "반: ______       이름: ______       점수: ______       선생님 확인: ______"
    info_p = dst.add_paragraph(info_text)
    set_runs_font(info_p, size_pt=12, bold=False)
    info_p.alignment = 1  # 가운데
    dst.add_paragraph("")

    # ---------- 본문 문제 ----------
    answer_map = {}
    next_blank_num = 1

    for para in src.paragraphs:
        orig_text = para.text.strip()
        if not orig_text:
            dst.add_paragraph("")
            continue

        tokens = tokenize_preserve_spacing(orig_text)
        try:
            tagged = pos_tag(tokens)
        except Exception:
            tagged = [(t, 'NN') for t in tokens]

        candidate_indices = []
        for i, (tok, tg) in enumerate(tagged):
            if is_candidate_token(tok):
                if pos_choice == "전체":
                    candidate_indices.append(i)
                else:
                    if tg in POS_GROUPS.get(pos_choice, set()):
                        candidate_indices.append(i)

        if not candidate_indices:
            candidate_indices = [i for i, (tok, tg) in enumerate(tagged) if is_candidate_token(tok)]

        n_candidates = len(candidate_indices)
        n_blanks = max(0, int(round(n_candidates * blank_ratio_fraction)))
        n_blanks = min(n_blanks, n_candidates)

        chosen = []
        if n_blanks > 0 and n_candidates > 0:
            chosen = random.sample(candidate_indices, n_blanks)

        out_tokens = list(tokens)
        for idx in sorted(chosen):
            original_word = tokens[idx]
            underline = "_" * max(3, len(original_word))
            out_tokens[idx] = f"({next_blank_num}){underline}"
            answer_map[next_blank_num] = original_word
            next_blank_num += 1

        para_text = assemble_tokens(out_tokens)
        p = dst.add_paragraph(para_text)
        set_runs_font(p, size_pt=11)

    # ---------- 정답지 ----------
    dst.add_page_break()
    title = dst.add_paragraph("📝 정답지 (Answer Sheet)")
    set_runs_font(title, size_pt=13, bold=True)

    total_answers = len(answer_map)
    if total_answers == 0:
        dst.add_paragraph("정답 항목이 없습니다.")
    else:
        num_cols = 3
        rows_needed = math.ceil(total_answers / num_cols)
        answers_table = dst.add_table(rows=rows_needed, cols=num_cols)
        answers_table.style = "Table Grid"

        for col in range(num_cols):
            for row in range(rows_needed):
                idx = col * rows_needed + row + 1
                cell = answers_table.cell(row, col)
                if idx <= total_answers:
                    cell.text = f"{idx}. {answer_map[idx]}"

    out = BytesIO()
    dst.save(out)
    out.seek(0)
    return out

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Blank Test Generator", layout="wide")
st.title("📘 Blank Test Generator")
st.markdown("업로드한 Word(.docx)에서 특정 품사만 선택하여 랜덤으로 빈칸을 생성하고, 마지막 페이지에 정답지를 자동으로 만들어 줍니다.")

# 설정
pos_choice = st.selectbox("빈칸으로 만들 품사 선택", ["전체", "동사", "명사", "형용사", "부사"])
blank_pct = st.slider("빈칸 비율 (%)", min_value=5, max_value=80, value=20, step=5)

uploaded_file = st.file_uploader("Word(.docx) 파일 업로드", type=["docx"])

if uploaded_file is not None:
    if st.button("시험지 생성 및 다운로드"):
        try:
            uploaded_file.seek(0)
            out = process_docx_with_answer(uploaded_file, pos_choice, blank_pct / 100.0)
            st.success("시험지가 생성되었습니다.")

            # 파일 이름 자동 생성
            original_name = uploaded_file.name
            base_name = os.path.splitext(original_name)[0]
            final_file_name = f"{base_name}_빈칸시험지+답지.docx"

            st.download_button(
                label="⬇️ 시험지(.docx) 다운로드 (문제 + 정답지 포함)",
                data=out,
                file_name=final_file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error("시험지 생성 중 오류가 발생했습니다.")
            st.exception(e)
else:
    st.info("먼저 Word(.docx) 파일을 업로드하세요.")
