import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import random
from io import BytesIO
import nltk

nltk.download('punkt')

# --- 페이지 설정 ---
st.set_page_config(page_title="Your Blank Test Generator", layout="wide")
st.title("📝 Your Blank Test Generator")
st.markdown("""
워드 파일(.docx)에서 텍스트를 불러와 **랜덤 단어를 빈칸 처리**하고,  
실제 시험지 형식으로 **자동 답지**까지 생성하는 앱입니다.
""")

# --- 파일 업로드 ---
uploaded_file = st.file_uploader("📂 워드 파일 업로드 (.docx)", type=["docx"])

# --- 빈칸 비율 ---
blank_ratio = st.slider("빈칸 비율 (%)", min_value=10, max_value=90, value=25, step=5)

# --- 함수: 랜덤 빈칸 생성 ---
def generate_random_blank_text(text, ratio):
    words = nltk.word_tokenize(text)
    n_blanks = max(1, int(len(words) * ratio / 100))
    blanks = {}
    
    if len(words) > 0:
        blank_indices = random.sample(range(len(words)), min(n_blanks, len(words)))
        for idx in blank_indices:
            blanks[idx] = words[idx]
            words[idx] = "_" * len(words[idx])
    return ' '.join(words), blanks

# --- 함수: 테두리 설정 ---
def set_paragraph_border(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for border_name in ['top','left','bottom','right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '4')
        border.set(qn('w:color'), '000000')
        pBdr.append(border)
    pPr.append(pBdr)

# --- 함수: 시험지 생성 ---
def process_docx_with_answer(file, ratio):
    doc = Document(file)
    new_doc = Document()
    
    # 기본 여백 설정
    sections = new_doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    all_answers = []  # 답지 저장
    
    # --- 시험지 상단 정보 ---
    header_table = new_doc.add_table(rows=2, cols=4)
    header_table.autofit = True
    header_table.style = 'Table Grid'
    
    cells = header_table.rows[0].cells
    cells[0].text = "반:"
    cells[1].text = ""
    cells[2].text = "이름:"
    cells[3].text = ""
    
    cells2 = header_table.rows[1].cells
    cells2[0].text = "점수:"
    cells2[1].text = ""
    cells2[2].text = "선생님 확인:"
    cells2[3].text = ""
    
    for row in header_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.font.size = Pt(12)
    
    new_doc.add_paragraph("")  # 줄간격

    # --- 본문 문단 처리 ---
    for para in doc.paragraphs:
        if para.text.strip() != "":
            blank_para, blanks = generate_random_blank_text(para.text, ratio)
            p = new_doc.add_paragraph(blank_para)
            set_paragraph_border(p)
            
            # 답지에 추가
            if blanks:
                all_answers.append({'text': para.text, 'blanks': blanks})
    
    # --- 답지 페이지 ---
    new_doc.add_page_break()
    answer_title = new_doc.add_paragraph("📝 정답지 (Answer Sheet)")
    answer_title.bold = True
    
    for i, item in enumerate(all_answers, 1):
        answer_line = f"{i}. "
        sorted_indices = sorted(item['blanks'].keys())
        for idx in sorted_indices:
            answer_line += item['blanks'][idx] + "  "
        new_doc.add_paragraph(answer_line.strip())
    
    # 메모리 상 저장
    output = BytesIO()
    new_doc.save(output)
    output.seek(0)
    return output

# --- 결과 처리 ---
if uploaded_file:
    st.success("✅ 파일 업로드 완료!")
    output_file = process_docx_with_answer(uploaded_file, blank_ratio)
    
    st.markdown("### 다운로드")
    st.download_button(
        label="⬇️ 시험지 + 답지 다운로드",
        data=output_file,
        file_name="random_blank_test_with_answer.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# --- 푸터 ---
st.markdown("---")
st.markdown("Made with ❤️ by Your Blank Test Generator")
