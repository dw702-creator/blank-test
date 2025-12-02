import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import nltk
from nltk import pos_tag, word_tokenize
import random
import re
import math

# ---------- NLTK data (필요 시 자동 다운로드) ----------
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

try:
    nltk.data.find('taggers/averaged_perceptron_tagger')
except LookupError:
    nltk.download('averaged_perceptron_tagger', quiet=True)

# ---------- POS 그룹 매핑 ----------
POS_GROUPS = {
    "동사": {"VB", "VBD", "VBG", "VBN", "VBP", "VBZ"},
    "명사": {"NN", "NNS", "NNP", "NNPS"},
    "형용사": {"JJ", "JJR", "JJS"},
    "부사": {"RB", "RBR", "RBS"},
}

# ---------- 헬퍼: 토큰 후보 판단 ----------
TOKEN_CANDIDATE_RE = re.compile(r"[A-Za-z0-9\uac00-\ud7a3]+")  # 알파벳, 숫자, 한글 포함

def is_candidate_token(tok):
    return bool(TOKEN_CANDIDATE_RE.search(tok))

# ---------- 헬퍼: 문장 토큰화 및 재조립 ----------
def tokenize_preserve_spacing(text):
    """
    word_tokenize로 토큰화 후, punctuation 붙임 규칙을 적용해 다시 문자열을 조립하기 쉬운 토큰 리스트를 반환.
    반환: tokens (list)
    """
    tokens = word_tokenize(text)
    return tokens

def assemble_tokens(tokens):
    """
    토큰 리스트를 문자열로 복원.
    punctuation(구두점) 앞에는 공백 없이 붙이고, 그 외에는 공백을 넣음.
    """
    out = ""
    for i, t in enumerate(tokens):
        if i == 0:
            out += t
            continue
        if re.fullmatch(r"[^\w\s]", t):  # punctuation
            out += t
        else:
            # 이전이 opening quote? (간단 처리) 항상 공백 추가
            out += " " + t
    return out

# ---------- 문단 테두리 적용 ----------
def set_paragraph_border(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    # 기존 pBdr 제거 (중복 방지)
    existing = pPr.find(qn('w:pBdr'))
    if existing is not None:
        pPr.remove(existing)
    pBdr = OxmlElement('w:pBdr')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '4')
        border.set(qn('w:color'), '000000')
        pBdr.append(border)
    pPr.append(pBdr)

# ---------- 폰트 설정 (run 단위) ----------
def set_runs_font(paragraph, size_pt=11, bold=False):
    for run in paragraph.runs:
        run.font.size = Pt(size_pt)
        run.font.bold = bold

def set_cell_font(cell, size_pt=11, bold=False):
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size_pt)
            r.font.bold = bold

# ---------- 핵심: 문서 생성 함수 ----------
def process_docx_with_answer(file_like, pos_choice, blank_ratio_fraction):
    """
    file_like: 업로드된 .docx 파일 객체
    pos_choice: "전체" 또는 "동사"/"명사"/"형용사"/"부사"
    blank_ratio_fraction: 0~1 사이 (예: 0.2)
    """
    src = Document(file_like)
    dst = Document()

    # 여백 설정
    for section in dst.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # 상단 헤더 테이블 (2x4)
    header = dst.add_table(rows=2, cols=4)
    header.style = 'Table Grid'
    header.autofit = True

    # 첫 줄 라벨들
    header.cell(0,0).text = "반"
    header.cell(0,1).text = ""
    header.cell(0,2).text = "이름"
    header.cell(0,3).text = ""
    # 둘째 줄
    header.cell(1,0).text = "점수"
    header.cell(1,1).text = ""
    header.cell(1,2).text = "선생님 확인"
    header.cell(1,3).text = ""

    # 폰트 조정
    for r in header.rows:
        for c in r.cells:
            set_cell_font(c, size_pt=11, bold=True)

    dst.add_paragraph("")  # 간격

    # 전체 텍스트를 문단 단위로 순회하며 문제 생성
    answer_map = {}   # { 번호: word }
    next_blank_num = 1

    for para in src.paragraphs:
        orig_text = para.text.strip()
        if not orig_text:
            dst.add_paragraph("")
            continue

        tokens = tokenize_preserve_spacing(orig_text)
        # POS 태깅 (pos_tag expects list of tokens)
        try:
            tagged = pos_tag(tokens)
        except Exception:
            # 만약 오류가 나면 간단 fallback: 모든 토큰에 'NN' 부여
            tagged = [(t, 'NN') for t in tokens]

        # 후보 인덱스: POS가 선택된 그룹에 속하고 토큰이 알파벳/숫자/한글 포함
        candidate_indices = []
        for i, (tok, tg) in enumerate(tagged):
            if is_candidate_token(tok):
                if pos_choice == "전체":
                    candidate_indices.append(i)
                else:
                    if tg in POS_GROUPS.get(pos_choice, set()):
                        candidate_indices.append(i)

        # 만약 선택된 품사가 문단에 하나도 없으면(후보 없음), 후보를 전체 단어로 확장
        if not candidate_indices:
            candidate_indices = [i for i, (tok, tg) in enumerate(tagged) if is_candidate_token(tok)]

        # 선택할 빈칸 수
        n_candidates = len(candidate_indices)
        n_blanks = max(0, int(round(n_candidates * blank_ratio_fraction)))  # 0 허용
        # 보장: n_blanks <= n_candidates
        n_blanks = min(n_blanks, n_candidates)

        chosen = []
        if n_blanks > 0 and n_candidates > 0:
            chosen = random.sample(candidate_indices, n_blanks)

        # 대체할 토큰 리스트 복사
        out_tokens = list(tokens)

        # 채우기: chosen 인덱스들을 번호 순으로 정렬하여 처리
        for idx in sorted(chosen):
            original_word = tokens[idx]
            # 언더바 길이: 원래 단어 길이 (유니코드 길이)
            underline = "_" * max(3, len(original_word))  # 최소 길이 3으로 표시 깔끔하게
            out_tokens[idx] = f"({next_blank_num}){underline}"
            answer_map[next_blank_num] = original_word
            next_blank_num += 1

        # assemble
        para_text = assemble_tokens(out_tokens)
        p = dst.add_paragraph(para_text)
        set_runs_font(p, size_pt=11, bold=False)
        set_paragraph_border(p)

    # --- 답지 페이지: 마지막 페이지에 추가 ---
    dst.add_page_break()
    title = dst.add_paragraph("📝 정답지 (Answer Sheet)")
    set_runs_font(title, size_pt=13, bold=True)

    total_answers = len(answer_map)
    if total_answers == 0:
        dst.add_paragraph("정답 항목이 없습니다.")
    else:
        # 3열 (columns) 구성, 열을 세로 방향으로 채우기
        num_cols = 3
        rows_needed = math.ceil(total_answers / num_cols)
        answers_table = dst.add_table(rows=rows_needed, cols=num_cols)
        answers_table.style = "Table Grid"

        # Fill column by column, top to bottom
        # mapping: for col in 0..num_cols-1:
        #   for row in 0..rows_needed-1:
        #       idx = col*rows_needed + row + 1
        for col in range(num_cols):
            for row in range(rows_needed):
                idx = col * rows_needed + row + 1
                cell = answers_table.cell(row, col)
                if idx <= total_answers:
                    cell.text = f"{idx}. {answer_map[idx]}"
                    set_cell_font(cell, size_pt=11, bold=False)
                else:
                    cell.text = ""
    # 메모리에 저장
    out = BytesIO()
    dst.save(out)
    out.seek(0)
    return out

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="연세영어학원 - 자동 빈칸 출제기", layout="wide")
st.title("📘 연세영어학원 자동 빈칸 출제기")
st.markdown("업로드한 Word(.docx)에서 특정 품사만 선택하여 랜덤으로 빈칸을 생성하고, 마지막 페이지에 정답지를 자동으로 만들어 줍니다.")

with st.sidebar:
    st.header("설정")
    pos_choice = st.selectbox("1) 빈칸으로 만들 품사 선택", ["전체", "동사", "명사", "형용사", "부사"])
    blank_pct = st.slider("2) 빈칸 비율 (%)", min_value=5, max_value=80, value=20, step=5,
                          help="선택된 품사 후보들 중에서 몇 %를 빈칸으로 만들지 결정합니다.")
    st.write("")
    st.markdown("⚠️ 한글 문서는 POS 태깅 정확도가 떨어질 수 있습니다.")
    preview_count = st.number_input("미리보기: 문단 수", min_value=1, max_value=20, value=5)

uploaded_file = st.file_uploader("Word(.docx) 파일 업로드", type=["docx"])

if uploaded_file is not None:
    st.success("파일 업로드 확인됨.")
    # 미리보기
    try:
        preview_doc = Document(uploaded_file)
        st.subheader("문서 미리보기 (최대 {} 문단)".format(preview_count))
        shown = 0
        for p in preview_doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            st.write(f"- {text}")
            shown += 1
            if shown >= preview_count:
                break
        uploaded_file.seek(0)
    except Exception as e:
        st.error("문서 미리보기 실패: 업로드된 파일이 올바른 docx 파일인지 확인하세요.")
        st.exception(e)

    if st.button("시험지 생성 및 다운로드"):
        try:
            uploaded_file.seek(0)
            out = process_docx_with_answer(uploaded_file, pos_choice, blank_pct / 100.0)
            st.success("시험지가 생성되었습니다. 아래 버튼으로 다운로드하세요.")
            st.download_button(
                label="⬇️ 시험지(.docx) 다운로드 (문제 + 정답지 포함)",
                data=out,
                file_name="blank_test_with_answer.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error("시험지 생성 중 오류가 발생했습니다.")
            st.exception(e)
else:
    st.info("먼저 Word(.docx) 파일을 업로드하세요.")
