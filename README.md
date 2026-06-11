# AI 그림책 질문수업 설계기

초보 교사가 학년, 주제, 그림책을 입력하면 질문, 활동, 활동지, 지도안, 평가, 학부모 안내문을 자동 생성하는 Streamlit 웹앱입니다.

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## OpenAI API Key 설정

로컬 실행 시 환경변수로 설정합니다.

```bash
export OPENAI_API_KEY="your_api_key"
```

Streamlit Community Cloud 배포 시:

1. GitHub에 이 폴더를 업로드합니다.
2. Streamlit Community Cloud에서 저장소를 연결합니다.
3. App secrets에 아래 내용을 입력합니다.

```toml
OPENAI_API_KEY="your_api_key"
```

## 주요 기능

- 질문 생성
- 활동 생성
- 활동지 생성
- 지도안 생성
- 평가 생성
- 학부모 안내문 생성
- Word/Markdown 다운로드
