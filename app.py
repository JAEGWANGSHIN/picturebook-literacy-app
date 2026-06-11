from __future__ import annotations

import os
from datetime import datetime
from io import BytesIO

import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt

st.set_page_config(
    page_title="AI 그림책 질문수업 설계기",
    page_icon="📚",
    layout="wide",
)

APP_TITLE = "📚 AI 그림책 질문수업 설계기"
APP_DESC = "학년, 주제, 그림책만 입력하면 초기 문해력 수업 초안을 자동으로 만들어 줍니다."

SYSTEM_PROMPT = """
당신은 초등 초기 문해력, 그림책 수업, 질문 중심 수업, 생성형 AI 활용 수업 설계 전문가입니다.
초보 교사가 바로 사용할 수 있도록 구체적이고 실제적인 수업안을 작성합니다.
반드시 한국어로 작성합니다.
수업은 그림책을 단순히 읽어주는 것이 아니라, 질문-대화-활동-성찰로 이어지게 설계합니다.
초기 문해력 요소는 음운인식, 어휘, 이야기 이해, 추론, 배경지식, 감정 이해, 표현 능력 중 적절한 것을 반영합니다.
질문은 사실 질문, 추론 질문, 평가 질문, 삶 연결 질문이 균형 있게 들어가야 합니다.
학생 발달 수준에 맞는 쉬운 언어를 사용합니다.
민감한 심리·상담 주제는 교사가 진단하지 않고 학생의 마음을 안전하게 표현하도록 돕는 방향으로 작성합니다.
"""

DEFAULT_BOOKS = [
    "알사탕", "강아지똥", "수박 수영장", "이상한 손님", "내 귀는 짝짝이",
    "너는 특별하단다", "중요한 사실", "고슴도치 X", "100만 번 산 고양이",
    "빈집에 온 손님", "무지개 물고기", "구름빵", "돼지책"
]

THEMES = [
    "자존감", "친구 관계", "감정 이해", "감정 조절", "다양성 존중", "정체성",
    "상상력", "가족", "두려움", "용기", "배려", "의사소통", "어휘 확장", "이야기 이해", "추론하기"
]


def get_client() -> OpenAI:
    api_key = st.secrets.get("OPENAI_API_KEY", None) or os.getenv("OPENAI_API_KEY")
    if not api_key:
        st.error("OpenAI API 키가 설정되지 않았습니다. Streamlit Secrets에 OPENAI_API_KEY를 등록해 주세요.")
        st.stop()
    return OpenAI(api_key=api_key)


def build_prompt(grade: str, theme: str, book: str, lesson_time: str, student_context: str, output_depth: str) -> str:
    return f"""
다음 조건을 바탕으로 'AI 그림책 질문수업 설계안'을 작성해 주세요.

[입력 조건]
- 학년: {grade}
- 수업 주제: {theme}
- 그림책: {book}
- 수업 시간: {lesson_time}
- 학생 특성: {student_context if student_context else '특별한 조건 없음'}
- 결과 상세 수준: {output_depth}

[출력 형식]
아래 항목을 반드시 모두 포함해 주세요.

1. 수업 개요
- 수업명
- 대상 학년
- 그림책
- 핵심 주제
- 초기 문해력 요소
- 수업 목표 3개

2. 그림책 활용 포인트
- 이 그림책이 해당 주제에 적합한 이유
- 학생들이 읽어야 할 글 요소
- 학생들이 읽어야 할 그림 요소
- 교사가 주의해야 할 점

3. 질문 생성
- 읽기 전 질문 3개
- 읽는 중 질문 5개
- 읽은 후 질문 5개
- 사실 질문, 추론 질문, 평가 질문, 삶 연결 질문으로 구분 표시

4. 활동 생성
- 활동 1: 도입 활동
- 활동 2: 중심 활동
- 활동 3: 표현/정리 활동
각 활동마다 활동 목표, 준비물, 진행 방법, 교사 발문, 예상 학생 반응을 포함

5. 활동지 초안
- 학생용 활동지 제목
- 안내 문장
- 문항 5개
- 그림 또는 글쓰기 활동 1개

6. 지도안
- 도입 / 전개 / 정리
- 시간 배분
- 교사 발문
- 학생 활동
- 자료 및 유의점

7. 평가
- 관찰 평가 기준 4개
- 학생 자기평가 문항 3개
- 교사용 피드백 문장 예시 5개

8. 학부모 안내문
- 가정에 보내는 안내문 형식
- 오늘 읽은 그림책과 수업 주제 소개
- 가정 대화 질문 3개
- 따뜻하고 전문적인 어조

9. AI 활용 팁
- 이 수업을 준비할 때 교사가 AI에 추가로 물어볼 수 있는 프롬프트 5개
"""


def generate_lesson(prompt: str, model: str) -> str:
    client = get_client()
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        temperature=0.7,
    )
    return response.choices[0].message.content or ""


def make_docx(text: str, title: str) -> bytes:
    doc = Document()
    styles = doc.styles
    styles["Normal"].font.name = "맑은 고딕"
    styles["Normal"].font.size = Pt(10.5)

    doc.add_heading(title, level=1)
    doc.add_paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
        elif stripped.startswith("# "):
            doc.add_heading(stripped.replace("# ", ""), level=1)
        elif stripped.startswith("## "):
            doc.add_heading(stripped.replace("## ", ""), level=2)
        elif stripped.startswith("### "):
            doc.add_heading(stripped.replace("### ", ""), level=3)
        else:
            doc.add_paragraph(stripped)

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def render_header() -> None:
    st.markdown(
        """
        <style>
        .main-title {font-size: 2.4rem; font-weight: 800; margin-bottom: 0.2rem;}
        .sub-text {font-size: 1.05rem; color: #555; margin-bottom: 1.2rem;}
        .info-box {background:#fff7e6; padding:1rem; border-radius:14px; border:1px solid #ffe0a3;}
        .result-box {background:#f8fafc; padding:1rem; border-radius:14px; border:1px solid #e2e8f0;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f"<div class='main-title'>{APP_TITLE}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='sub-text'>{APP_DESC}</div>", unsafe_allow_html=True)


def main() -> None:
    render_header()

    with st.sidebar:
        st.header("⚙️ 설정")
        model = st.selectbox(
            "AI 모델",
            ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini", "gpt-4.1"],
            index=0,
        )
        output_depth = st.radio("결과 상세 수준", ["간단", "보통", "상세"], index=1)
        st.divider()
        st.caption("API 키는 Streamlit Secrets에 OPENAI_API_KEY로 저장합니다.")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1️⃣ 수업 조건 입력")
        grade = st.selectbox("학년", ["유치원", "초등 1학년", "초등 2학년", "초등 3학년", "초등 4학년", "초등 5학년", "초등 6학년"])
        theme = st.selectbox("수업 주제", THEMES)
        book_mode = st.radio("그림책 입력 방식", ["목록에서 선택", "직접 입력"], horizontal=True)
        if book_mode == "목록에서 선택":
            book = st.selectbox("그림책", DEFAULT_BOOKS)
        else:
            book = st.text_input("그림책 제목", placeholder="예: 알사탕")

        lesson_time = st.selectbox("수업 시간", ["40분", "80분", "120분", "프로젝트 3차시", "프로젝트 5차시"])
        student_context = st.text_area(
            "학생 특성 또는 수업 상황",
            placeholder="예: 1학년 입학 초기, 친구 관계가 아직 서툰 편, 글쓰기 부담이 큼 등",
            height=120,
        )

        generate_btn = st.button("✨ 수업 설계안 생성하기", type="primary", use_container_width=True)

    with col2:
        st.subheader("2️⃣ 생성될 결과")
        st.markdown(
            """
            <div class='info-box'>
            ✅ 질문 생성<br>
            ✅ 활동 생성<br>
            ✅ 활동지 생성<br>
            ✅ 지도안 생성<br>
            ✅ 평가 생성<br>
            ✅ 학부모 안내문 생성<br>
            ✅ AI 추가 프롬프트 생성
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info("초보 선생님이 바로 수정해서 쓸 수 있는 형태로 생성됩니다.")

    if generate_btn:
        if not book.strip():
            st.warning("그림책 제목을 입력해 주세요.")
            st.stop()

        prompt = build_prompt(grade, theme, book, lesson_time, student_context, output_depth)
        with st.spinner("AI가 그림책 질문수업 설계안을 만드는 중입니다..."):
            try:
                result = generate_lesson(prompt, model)
            except Exception as e:
                st.error(f"생성 중 오류가 발생했습니다: {e}")
                st.stop()

        st.session_state["result"] = result
        st.session_state["title"] = f"{book}_{theme}_질문수업설계안"

    if "result" in st.session_state:
        st.divider()
        st.subheader("3️⃣ 생성 결과")
        st.markdown(st.session_state["result"])

        docx_bytes = make_docx(st.session_state["result"], st.session_state["title"])
        st.download_button(
            "📄 Word 파일로 다운로드",
            data=docx_bytes,
            file_name=f"{st.session_state['title']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.download_button(
            "📝 Markdown 파일로 다운로드",
            data=st.session_state["result"].encode("utf-8"),
            file_name=f"{st.session_state['title']}.md",
            mime="text/markdown",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
