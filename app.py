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
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── DB ──────────────────────────────────────────────────────────────
PICTUREBOOK_DB = [
    {"id":"pb001","title":"말놀이 동시집","author":"최승호·방시혁","theme":["음운인식","어휘"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"동물 이름과 의성어로 구성된 말놀이 동시 모음","literacy_elements":["음운인식","어휘"],"reason":"한국어 음소·음절 인식, 운율 체험","questions":{"사실":["이 시에서 어떤 동물들이 나오나요?"],"추론":["왜 동물 이름과 소리를 함께 썼을까요?"],"평가":["가장 재미있는 말놀이는 무엇인가요?"],"감정":["이 시를 읽을 때 어떤 기분이 드나요?"],"작가":["작가는 왜 말놀이로 시를 썼을까요?"],"삶연결":["내가 좋아하는 말놀이가 있나요?"]}},
    {"id":"pb002","title":"수수께끼야 놀자","author":"이상교","theme":["음운인식","어휘"],"grade":["유치원","초등 1학년"],"summary":"수수께끼 형식으로 의미와 소리를 연결","literacy_elements":["음운인식"],"reason":"소리 단서로 단어 맞추기; 파닉스 연결","questions":{"사실":["어떤 수수께끼들이 나오나요?"],"추론":["수수께끼 단서에서 소리가 어떤 역할을 하나요?"],"평가":["가장 어려운 수수께끼는 무엇이었나요?"],"감정":["수수께끼를 맞혔을 때 기분이 어떤가요?"],"작가":["작가는 왜 수수께끼 형식을 선택했을까요?"],"삶연결":["내가 만든 수수께끼를 친구에게 낼 수 있나요?"]}},
    {"id":"pb003","title":"Brown Bear, Brown Bear","author":"Bill Martin Jr.","theme":["음운인식","어휘"],"grade":["유치원","초등 1학년"],"summary":"색깔과 동물 이름이 반복되는 패턴 그림책","literacy_elements":["음운인식","어휘"],"reason":"반복 패턴으로 운율·예측 읽기","questions":{"사실":["갈색 곰이 보는 것은 무엇인가요?"],"추론":["왜 같은 패턴이 계속 반복될까요?"],"평가":["이 책에서 가장 마음에 드는 동물은?"],"감정":["반복되는 패턴을 읽을 때 어떤 느낌이 드나요?"],"작가":["작가는 왜 모든 동물에게 색깔을 붙여줬을까요?"],"삶연결":["내가 좋아하는 색깔과 동물을 연결한다면?"]}},
    {"id":"pb004","title":"단어 수집가","author":"Peter H. Reynolds","theme":["어휘","정체성"],"grade":["초등 1학년","초등 2학년","초등 3학년"],"summary":"소년이 세상의 단어들을 수집하는 이야기","literacy_elements":["어휘","음운인식"],"reason":"단어 가치 인식; 어휘 수집 동기 부여","questions":{"사실":["소년은 무엇을 수집하나요?"],"추론":["왜 단어를 모으는 것이 특별할까요?"],"평가":["내가 수집하고 싶은 단어 3개는?"],"감정":["좋아하는 단어를 발견했을 때 기분이 어떨까요?"],"작가":["작가는 왜 주인공이 단어를 수집하게 했을까요?"],"삶연결":["내가 가장 좋아하는 단어가 있나요?"]}},
    {"id":"pb005","title":"알사탕","author":"백희나","theme":["어휘","감정 이해","가족"],"grade":["초등 1학년","초등 2학년"],"summary":"신비한 사탕을 먹으면 주변의 소리가 들린다","literacy_elements":["어휘","감정"],"reason":"감각어·감정 어휘 풍부; 서정적 표현","questions":{"사실":["동동이는 어떤 사탕을 먹었나요?"],"추론":["할머니의 목소리는 왜 가장 크게 들렸을까요?"],"평가":["알사탕이 진짜 있다면, 누구의 말을 듣고 싶나요?"],"감정":["아빠의 '사랑한다'는 말을 들었을 때 동동이는?"],"작가":["작가는 왜 동동이가 혼자 논다는 내용을 먼저 알려줬을까요?"],"삶연결":["가족의 마음속 이야기를 들을 수 있다면?"]}},
    {"id":"pb006","title":"구름빵","author":"백희나","theme":["어휘","가족","상상력"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"비 오는 날 구름으로 빵을 만들어 날아다니는 이야기","literacy_elements":["어휘","이야기이해"],"reason":"요리·자연 어휘; 판타지 어휘 확장","questions":{"사실":["구름빵을 먹으면 어떤 일이 일어나나요?"],"추론":["엄마가 아이들에게 빵을 만들어 준 이유는?"],"평가":["구름으로 빵 말고 무엇을 만들고 싶나요?"],"감정":["아빠를 위해 날아갈 때 아이들은 어떤 마음이었을까요?"],"작가":["작가는 왜 비 오는 날을 배경으로 설정했을까요?"],"삶연결":["가족을 위해 내가 할 수 있는 작은 일은?"]}},
    {"id":"pb007","title":"수박 수영장","author":"안녕달","theme":["어휘","배경지식","다양성 존중","상상력"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"커다란 수박 속 수영장에서 동네 사람들이 함께 노는 상상","literacy_elements":["어휘","배경지식"],"reason":"여름·공동체·감각 어휘; 배경지식 확장","questions":{"사실":["수박 수영장에는 어떤 사람들이 놀러 왔나요?"],"추론":["이 수영장은 꿈인가요, 현실인가요?"],"평가":["모두 함께 노는 것이 왜 중요할까요?"],"감정":["처음 수박 수영장을 발견했을 때 기분은?"],"작가":["작가는 왜 '수박'으로 수영장을 만들었을까요?"],"삶연결":["내가 상상하는 가장 멋진 공간은?"]}},
    {"id":"pb008","title":"이상한 손님","author":"백희나","theme":["어휘","추론하기","배려"],"grade":["초등 1학년","초등 2학년"],"summary":"비 오는 날 하늘 나라에서 길 잃은 아이가 찾아온 이야기","literacy_elements":["어휘","추론"],"reason":"감정 어휘·비유 표현; 상황 맥락 어휘 추론","questions":{"사실":["두 아이는 이상한 손님을 어떻게 돌봐줬나요?"],"추론":["이상한 손님은 어디서 온 누구일까요?"],"평가":["낯선 손님을 돕는 것이 옳은 일일까요?"],"감정":["손님이 떠났을 때 두 아이의 마음은?"],"작가":["작가는 왜 손님의 정체를 끝까지 숨겼을까요?"],"삶연결":["예상치 못한 손님이 찾아온다면 어떻게 대접하고 싶나요?"]}},
    {"id":"pb009","title":"내 귀는 짝짝이","author":"율리 슈타르크","theme":["자존감","다양성 존중","정체성"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"귀가 다른 토끼가 자신을 있는 그대로 받아들이는 이야기","literacy_elements":["감정","자존감"],"reason":"신체 다양성 수용; 자기 긍정 감정 어휘","questions":{"사실":["주인공의 귀는 어떻게 생겼나요?"],"추론":["주인공이 처음에 자신의 귀를 부끄러워한 이유는?"],"평가":["남과 다르다는 것이 나쁜 일일까요?"],"감정":["다른 사람들이 나를 이상하게 볼 때 어떤 기분인가요?"],"작가":["작가는 왜 주인공을 토끼로 만들었을까요?"],"삶연결":["내가 다른 사람과 다른 특별한 점은?"]}},
    {"id":"pb010","title":"너는 특별하단다","author":"맥스 루케이도","theme":["자존감","정체성"],"grade":["유치원","초등 1학년","초등 2학년","초등 3학년"],"summary":"작은 나무 사람 펀치넬로가 자신의 가치를 깨닫는 이야기","literacy_elements":["감정","이야기이해"],"reason":"자존감 언어; 평가와 자기 가치 탐색","questions":{"사실":["펀치넬로는 어떤 딱지를 받았나요?"],"추론":["나무 조각가 엘리는 누구를 상징할까요?"],"평가":["다른 사람의 평가가 나의 가치를 결정할 수 있을까요?"],"감정":["나쁜 딱지를 받았을 때 펀치넬로는 어떤 기분이었을까요?"],"작가":["작가는 왜 나무 인형 이야기로 이 주제를 전달했을까요?"],"삶연결":["다른 사람의 말에 상처받은 적이 있나요?"]}},
    {"id":"pb011","title":"중요한 사실","author":"마가렛 와이즈 브라운","theme":["정체성","어휘"],"grade":["초등 1학년","초등 2학년"],"summary":"사물의 가장 중요한 특성에 대해 이야기하는 철학적 그림책","literacy_elements":["어휘","추론"],"reason":"핵심 특성 파악; 자아에 대한 질문 생성","questions":{"사실":["각 사물의 '중요한 사실'은 무엇인가요?"],"추론":["왜 한 가지 사실이 가장 중요하다고 했을까요?"],"평가":["나에 대한 '중요한 사실'은 무엇일까요?"],"감정":["자신에 대해 가장 중요한 사실을 말할 때 어떤 기분인가요?"],"작가":["작가는 왜 평범한 사물을 선택했을까요?"],"삶연결":["내 친구에 대한 '중요한 사실'은?"]}},
    {"id":"pb012","title":"고슴도치 X","author":"에밀리 그레이벳","theme":["친구 관계","감정 이해","의사소통"],"grade":["초등 1학년","초등 2학년"],"summary":"편지를 쓰다가 계속 틀려서 X로 지워나가는 고슴도치 이야기","literacy_elements":["감정","이야기이해"],"reason":"감정 표현의 어려움; 쓰기와 감정 연결","questions":{"사실":["고슴도치는 누구에게 편지를 쓰고 있나요?"],"추론":["고슴도치가 말하고 싶었던 것은?"],"평가":["마음을 전할 때 말과 글 중 어느 것이 더 쉬운가요?"],"감정":["고슴도치는 어떤 마음이었을까요?"],"작가":["작가는 왜 고슴도치를 주인공으로 선택했을까요?"],"삶연결":["마음을 전하기 어려웠던 경험이 있나요?"]}},
    {"id":"pb013","title":"100만 번 산 고양이","author":"사노 요코","theme":["자존감","정체성","감정 이해"],"grade":["초등 2학년","초등 3학년","초등 4학년"],"summary":"100만 번을 살면서 진정한 사랑을 깨달은 고양이 이야기","literacy_elements":["이야기이해","감정"],"reason":"삶과 사랑의 의미; 감정 변화 추적","questions":{"사실":["고양이가 처음으로 운 이유는?"],"추론":["왜 100만 번 살았을 때는 죽지 않았을까요?"],"평가":["진정한 사랑이란 무엇일까요?"],"감정":["흰 고양이와 함께할 때 행복했던 이유는?"],"작가":["작가는 왜 이 이야기를 어린이 책으로 썼을까요?"],"삶연결":["내가 가장 소중히 여기는 것은?"]}},
    {"id":"pb014","title":"빈집에 온 손님","author":"김유경","theme":["두려움","용기","상상력"],"grade":["초등 1학년","초등 2학년"],"summary":"홀로 집을 지키던 아이가 상상 속 손님을 맞이하는 이야기","literacy_elements":["추론","감정"],"reason":"두려움과 상상력; 감정 탐색","questions":{"사실":["어떤 손님이 찾아왔나요?"],"추론":["손님은 진짜인가요, 상상인가요?"],"평가":["혼자 있을 때의 두려움을 어떻게 극복할 수 있을까요?"],"감정":["혼자 집을 지킬 때 어떤 기분이었을까요?"],"작가":["작가는 왜 이런 이야기를 썼을까요?"],"삶연결":["혼자 있을 때 무서웠던 적이 있나요?"]}},
    {"id":"pb015","title":"무지개 물고기","author":"마르쿠스 피스터","theme":["어휘","배려","친구 관계"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"아름다운 비늘을 나눠주며 친구를 사귀는 물고기 이야기","literacy_elements":["어휘","감정"],"reason":"바닷속 어휘; 나눔·감정 표현 어휘","questions":{"사실":["무지개 물고기의 특별한 점은?"],"추론":["처음에 왜 비늘을 나눠주지 않았을까요?"],"평가":["특별한 것을 나눠주면 특별함이 줄어들까요?"],"감정":["아무도 친구가 되어주지 않을 때 어떤 마음이었을까요?"],"작가":["작가는 왜 비늘에 홀로그램을 사용했을까요?"],"삶연결":["내가 가진 것 중 나누면 더 기쁜 것은?"]}},
    {"id":"pb016","title":"강아지똥","author":"권정생","theme":["자존감","감정 이해","배려"],"grade":["초등 1학년","초등 2학년","초등 3학년"],"summary":"아무도 거들떠보지 않던 강아지똥이 민들레의 거름이 된다","literacy_elements":["어휘","감정"],"reason":"자연 어휘·존재 가치 어휘; 문학적 표현","questions":{"사실":["민들레는 강아지똥에게 무엇을 부탁했나요?"],"추론":["강아지똥이 처음에 슬펐던 이유는?"],"평가":["강아지똥은 정말 쓸모없는 존재였을까요?"],"감정":["아무도 필요 없다고 했을 때 강아지똥의 마음은?"],"작가":["작가는 왜 '강아지똥'을 주인공으로 선택했을까요?"],"삶연결":["처음엔 별 볼 일 없어 보였지만 나중에 소중했던 것이 있나요?"]}},
    {"id":"pb017","title":"돼지책","author":"앤서니 브라운","theme":["가족","다양성 존중","의사소통"],"grade":["초등 2학년","초등 3학년","초등 4학년"],"summary":"혼자 집안일을 하던 엄마가 집을 나가고 가족이 돼지로 변한다","literacy_elements":["이야기이해","추론"],"reason":"인물 동기·감정 변화 추적; 시각 상징 분석","questions":{"사실":["엄마가 집을 나간 후 가족들에게 어떤 일이 일어났나요?"],"추론":["엄마는 왜 집을 나갔을까요?"],"평가":["집안일은 누가 해야 할까요?"],"감정":["엄마가 없어졌을 때 두 아들은 어떤 마음이었을까요?"],"작가":["작가는 왜 가족이 돼지로 변하는 그림을 그렸을까요?"],"삶연결":["우리 집에서 나는 어떤 일을 돕고 있나요?"]}},
    {"id":"pb018","title":"꽃들에게 희망을","author":"트리나 폴러스","theme":["자존감","정체성","용기"],"grade":["초등 2학년","초등 3학년","초등 4학년"],"summary":"애벌레가 자아를 찾아 나비가 되는 우화","literacy_elements":["어휘","추론"],"reason":"삶의 의미 어휘; 변화·희망 주제 어휘","questions":{"사실":["호랑 애벌레는 처음에 무엇을 찾아 떠났나요?"],"추론":["기둥을 오르는 애벌레들은 무엇을 상징할까요?"],"평가":["진짜 원하는 것을 위해 모든 것을 포기하는 것이 옳을까요?"],"감정":["고치 속에 혼자 있을 때 어떤 마음이었을까요?"],"작가":["작가는 왜 나비 이야기를 애벌레 기둥으로 표현했을까요?"],"삶연결":["포기하고 싶었지만 끝까지 해낸 일이 있나요?"]}},
    {"id":"pb019","title":"고구마구마","author":"사이다","theme":["어휘","의사소통","감정 이해"],"grade":["유치원","초등 1학년"],"summary":"고구마가 '구마'라고만 말하는 반복 언어유희 그림책","literacy_elements":["어휘","음운인식"],"reason":"파닉스 연결; 반복 패턴 어휘 강화","questions":{"사실":["고구마는 뭐라고 말하나요?"],"추론":["'구마'라는 말이 각 상황에서 어떤 의미를 가질까요?"],"평가":["말이 없어도 마음을 전할 수 있을까요?"],"감정":["하고 싶은 말이 있는데 표현하기 어려웠던 적이 있나요?"],"작가":["작가는 왜 '구마'라는 한 단어만 반복했을까요?"],"삶연결":["말이 아닌 다른 방법으로 감정을 표현한 적이 있나요?"]}},
    {"id":"pb020","title":"괜찮아","author":"최숙희","theme":["자존감","다양성 존중","감정 이해"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"서로 다른 모습도 괜찮다는 자기 수용 이야기","literacy_elements":["감정","어휘"],"reason":"자기 수용·다양성 감정 표현 어휘","questions":{"사실":["이 책에서 어떤 동물들이 나오나요?"],"추론":["왜 자신과 다른 모습도 '괜찮다'고 하는 걸까요?"],"평가":["남과 다른 모습이 부끄러운 일인가요?"],"감정":["자신의 다른 점에 대해 처음에는 어떤 기분이었을까요?"],"작가":["작가는 왜 다양한 동물들로 이야기를 썼을까요?"],"삶연결":["내가 다른 사람과 달라서 걱정했던 점이 있나요?"]}},
    {"id":"pb021","title":"7년 동안의 잠","author":"박완서","theme":["배경지식","감정 이해"],"grade":["초등 2학년","초등 3학년"],"summary":"흉년 든 개미마을에 나타난 번데기를 둘러싼 이야기","literacy_elements":["이야기이해","배경지식"],"reason":"생태 배경지식; 발단·전개·결말 구조 분석","questions":{"사실":["개미들은 왜 번데기가 욕심났을까요?"],"추론":["'7년 동안의 잠'이 의미하는 것은?"],"평가":["배가 고프다면 무엇이든 먹어도 되나요?"],"감정":["번데기로 변해 긴 잠을 자는 매미의 기분은?"],"작가":["작가는 왜 이 이야기를 어린이를 위해 썼을까요?"],"삶연결":["기다려야만 이루어지는 것들을 생각해 본 적 있나요?"]}},
    {"id":"pb022","title":"나쁜 어린이 표","author":"황선미","theme":["자존감","감정 이해"],"grade":["초등 1학년","초등 2학년","초등 3학년"],"summary":"잘못을 저지른 어린이가 표를 붙이고 다니는 이야기","literacy_elements":["이야기이해","감정"],"reason":"원인-결과; 인물 내면 변화 이해","questions":{"사실":["어린이는 왜 나쁜 어린이 표를 받게 됐나요?"],"추론":["표가 없어지려면 어떻게 해야 할까요?"],"평가":["잘못을 했을 때 어떻게 하는 것이 좋을까요?"],"감정":["나쁜 어린이 표를 받았을 때 어떤 기분이었을까요?"],"작가":["작가는 왜 이런 이야기를 썼을까요?"],"삶연결":["잘못을 하고 용기를 내어 사과한 적이 있나요?"]}},
    {"id":"pb023","title":"선물","author":"이수지","theme":["상상력","감정 이해"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"눈이 오는 날의 감동을 글 없이 그림만으로 표현","literacy_elements":["추론"],"reason":"무자(wordless) 그림책; 그림 추론 최적","questions":{"사실":["그림 속에서 무슨 일이 일어나고 있나요?"],"추론":["선물이란 무엇을 의미할까요?"],"평가":["가장 좋은 선물은 어떤 것일까요?"],"감정":["선물을 받았을 때 어떤 기분이었을까요?"],"작가":["작가는 왜 글을 하나도 쓰지 않았을까요?"],"삶연결":["내가 받은 가장 소중한 선물은?"]}},
    {"id":"pb024","title":"지각대장 존","author":"John Burningham","theme":["상상력","의사소통"],"grade":["초등 1학년","초등 2학년"],"summary":"매일 지각하는 존의 기상천외한 이유","literacy_elements":["이야기이해","추론"],"reason":"사실과 상상 구별; 인물 관점 이해","questions":{"사실":["존은 왜 학교에 지각했나요?"],"추론":["존의 이야기는 진짜인가요, 상상인가요?"],"평가":["존이 솔직하게 말하는 게 좋을까요?"],"감정":["존은 학교 가는 것이 즐거웠을까요?"],"작가":["작가는 왜 어른들이 존의 말을 믿지 않게 썼을까요?"],"삶연결":["상상력을 발휘해서 재미있는 변명을 해본 적 있나요?"]}},
    {"id":"pb025","title":"100층짜리 집","author":"이와이 도시오","theme":["배경지식","상상력"],"grade":["유치원","초등 1학년"],"summary":"주인공이 100층까지 올라가며 여러 동물을 만나는 이야기","literacy_elements":["이야기이해","배경지식"],"reason":"순서·수 개념; 동물 생태 배경지식","questions":{"사실":["어떤 동물들이 몇 층에 살고 있나요?"],"추론":["각 층에 사는 동물들은 왜 그 높이를 선택했을까요?"],"평가":["나라면 몇 층에 살고 싶나요?"],"감정":["높이 올라가면서 어떤 기분이 들었을까요?"],"작가":["작가는 왜 100층이라는 설정을 선택했을까요?"],"삶연결":["100층짜리 집을 만든다면 내 층에는 무엇을 꾸미고 싶나요?"]}},
    {"id":"pb026","title":"으뜸 헤엄이(Swimmy)","author":"Leo Lionni","theme":["친구 관계","용기","배려"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"혼자 헤엄치는 물고기가 친구들과 힘을 합치는 이야기","literacy_elements":["이야기 재구성","감정"],"reason":"시각적 장면 순서로 재구성 용이","questions":{"사실":["스위미가 다른 물고기들에게 제안한 것은?"],"추론":["왜 스위미가 눈 역할을 맡았을까요?"],"평가":["함께 힘을 합치는 것이 왜 중요할까요?"],"감정":["혼자 남겨졌을 때 스위미는 어떤 기분이었을까요?"],"작가":["작가는 왜 검은색 물고기를 주인공으로 만들었을까요?"],"삶연결":["친구들과 힘을 합쳐서 해낸 경험이 있나요?"]}},
    {"id":"pb027","title":"The Very Hungry Caterpillar","author":"Eric Carle","theme":["배경지식","어휘"],"grade":["유치원","초등 1학년"],"summary":"배고픈 애벌레가 다양한 음식을 먹으며 성장하는 이야기","literacy_elements":["이야기 재구성","배경지식"],"reason":"요일·음식·변태 순서 재구성; 반복 구조","questions":{"사실":["애벌레는 무슨 요일에 무엇을 먹었나요?"],"추론":["왜 애벌레는 점점 더 많이 먹었을까요?"],"평가":["많이 먹는 것이 좋은 일일까요?"],"감정":["번데기 속에서 기다리는 동안 어떤 기분이었을까요?"],"작가":["작가는 왜 구멍 뚫린 책을 만들었을까요?"],"삶연결":["내가 제일 좋아하는 음식은?"]}},
    {"id":"pb028","title":"Where the Wild Things Are","author":"Maurice Sendak","theme":["감정 조절","상상력","가족"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"맥스가 상상의 세계로 여행하고 집으로 돌아오는 이야기","literacy_elements":["이야기 재구성","감정"],"reason":"여행 구조(출발-모험-귀환) 재구성 전형","questions":{"사실":["맥스는 왜 방에 갇히게 됐나요?"],"추론":["괴물들의 나라는 무엇을 상징할까요?"],"평가":["화가 날 때 어떻게 하는 것이 좋을까요?"],"감정":["맥스가 집으로 돌아왔을 때 따뜻한 저녁이 기다리고 있었던 이유는?"],"작가":["작가는 왜 맥스의 방이 점점 정글로 변하는 그림을 그렸을까요?"],"삶연결":["화가 났을 때 나만의 방법이 있나요?"]}},
    {"id":"pb029","title":"In My Heart","author":"Jo Witek","theme":["감정 이해","감정 조절"],"grade":["유치원","초등 1학년"],"summary":"다양한 감정을 신체 감각으로 묘사하는 그림책","literacy_elements":["감정","어휘"],"reason":"감정 어휘 10가지 명시적 학습","questions":{"사실":["이 책에서 어떤 감정들이 나오나요?"],"추론":["왜 감정마다 느껴지는 방식이 다를까요?"],"평가":["기쁨과 슬픔 중 어떤 감정이 더 중요할까요?"],"감정":["지금 내 마음속에는 어떤 감정이 있나요?"],"작가":["작가는 왜 감정을 신체 감각으로 설명했을까요?"],"삶연결":["오늘 어떤 감정을 느꼈나요?"]}},
    {"id":"pb030","title":"왜냐하면(Because)","author":"Mo Willems","theme":["상상력","배경지식"],"grade":["초등 1학년","초등 2학년"],"summary":"연쇄적 원인-결과로 이어지는 이야기","literacy_elements":["이야기이해"],"reason":"'왜?' 질문 구조를 시각적으로 보여줌","questions":{"사실":["이야기에서 어떤 연결이 이루어지나요?"],"추론":["작은 일 하나가 어떻게 큰 일로 이어질 수 있을까요?"],"평가":["우리 삶에서 '왜냐하면'으로 연결되는 것들이 있나요?"],"감정":["예상치 못한 결과가 생겼을 때 어떤 기분인가요?"],"작가":["작가는 왜 연쇄 연결 구조를 선택했을까요?"],"삶연결":["내가 한 작은 행동이 예상치 못한 결과로 이어진 적이 있나요?"]}},
    {"id":"pb031","title":"Two Bad Ants","author":"Chris Van Allsburg","theme":["상상력","용기"],"grade":["초등 2학년","초등 3학년"],"summary":"두 개미가 설탕 그릇으로 모험을 떠나는 이야기","literacy_elements":["추론"],"reason":"개미 시점으로 보는 세상: 시각 추론의 정수","questions":{"사실":["두 개미는 왜 무리에서 떨어졌나요?"],"추론":["개미들이 보는 컵과 토스터는 실제로 무엇인가요?"],"평가":["규칙을 어기는 것이 언제 문제가 될까요?"],"감정":["집으로 돌아왔을 때 개미들은 어떤 기분이었을까요?"],"작가":["작가는 왜 개미의 시점에서 이야기를 썼을까요?"],"삶연결":["아주 작은 시점으로 세상을 본다면 무엇이 달라 보일까요?"]}},
    {"id":"pb032","title":"Voices in the Park","author":"앤서니 브라운","theme":["다양성 존중","친구 관계","의사소통"],"grade":["초등 2학년","초등 3학년","초등 4학년"],"summary":"같은 공원 방문을 4명의 서로 다른 목소리로 이야기","literacy_elements":["추론","이야기이해"],"reason":"관점 추론; 같은 사건의 다른 해석","questions":{"사실":["몇 명의 목소리가 등장하나요?"],"추론":["왜 같은 사건을 다르게 기억할까요?"],"평가":["누구의 이야기가 '진짜'인가요?"],"감정":["각 인물은 공원에서 어떤 감정을 느꼈나요?"],"작가":["작가는 왜 4개의 다른 목소리를 만들었을까요?"],"삶연결":["같은 일을 나와 친구가 다르게 기억한 적이 있나요?"]}},
    {"id":"pb033","title":"선생님이 나를 모르면","author":"이상교","theme":["정체성","의사소통"],"grade":["초등 1학년"],"summary":"아이가 선생님에게 자신을 소개하는 이야기","literacy_elements":["이야기이해"],"reason":"나에 대한 질문 생성; 자기 이해 촉진","questions":{"사실":["아이는 선생님에게 무엇을 알려주고 싶었나요?"],"추론":["왜 선생님이 나를 아는 것이 중요할까요?"],"평가":["선생님에게 꼭 알려줘야 할 것은?"],"감정":["새 학년에 새 선생님을 만나면 어떤 기분인가요?"],"작가":["작가는 왜 아이의 목소리로 이야기를 썼을까요?"],"삶연결":["선생님이 나에 대해 꼭 알았으면 하는 것은?"]}},
    {"id":"pb034","title":"나는 어떻게 생겨났을까?","author":"과학그림책","theme":["배경지식","정체성"],"grade":["초등 1학년","초등 2학년"],"summary":"탄생의 과학적 사실을 어린이 눈높이로 설명","literacy_elements":["배경지식"],"reason":"배경지식 궁금증에서 질문 생성 자연 유도","questions":{"사실":["아기는 어떻게 태어나나요?"],"추론":["왜 모든 사람은 다르게 생겼을까요?"],"평가":["생명은 왜 소중한가요?"],"감정":["내가 태어났을 때 가족들은 어떤 기분이었을까요?"],"작가":["작가는 왜 이런 책을 어린이를 위해 썼을까요?"],"삶연결":["내가 태어난 날에 대해 들은 이야기가 있나요?"]}},
    {"id":"pb035","title":"The Invisible String","author":"Patrice Karst","theme":["감정 이해","두려움","가족"],"grade":["유치원","초등 1학년","초등 2학년"],"summary":"사랑하는 사람과의 보이지 않는 연결 이야기","literacy_elements":["감정"],"reason":"분리불안·연결감 감정; 저학년 적합","questions":{"사실":["보이지 않는 실은 무엇인가요?"],"추론":["왜 사랑하는 사람과의 연결은 눈에 보이지 않을까요?"],"평가":["멀리 있어도 마음이 연결될 수 있을까요?"],"감정":["가족이 보고 싶을 때 어떤 기분인가요?"],"작가":["작가는 왜 '실'을 연결의 상징으로 사용했을까요?"],"삶연결":["보고 싶은 사람이 생각날 때 어떻게 하나요?"]}},
]

def db_search(theme="", grade=""):
    return [b for b in PICTUREBOOK_DB if
            ((not theme) or any(theme in t for t in b["theme"])) and
            ((not grade) or grade in b["grade"])]

def db_get_by_title(title):
    return next((b for b in PICTUREBOOK_DB if b["title"] == title), None)

def db_all_themes():
    themes = set()
    for b in PICTUREBOOK_DB: themes.update(b["theme"])
    return sorted(themes)

# ── CSS ─────────────────────────────────────────────────────────────
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Gaegu:wght@700&family=Nanum+Gothic:wght@400;700;800&display=swap');

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > .main {
    background-color: #FFF9F0 !important;
}
[data-testid="stSidebar"] { display: none !important; }

.main .block-container {
    max-width: 760px !important;
    padding: 2rem 1.5rem 3rem !important;
}

/* ── 헤더 ── */
.app-header {
    text-align: center;
    padding: 2.2rem 1rem 1.6rem;
    margin-bottom: 0.5rem;
}
.app-icon {
    font-size: 2.8rem;
    display: block;
    margin-bottom: 0.5rem;
    animation: bob 3s ease-in-out infinite;
}
@keyframes bob { 0%,100%{transform:translateY(0);} 50%{transform:translateY(-6px);} }
.app-title {
    font-family: 'Gaegu', cursive !important;
    font-size: clamp(1.7rem, 5vw, 2.4rem) !important;
    color: #3D2B1F !important;
    line-height: 1.2 !important;
    margin: 0 0 0.4rem !important;
}
.app-sub {
    font-size: 0.92rem;
    color: #7D5A4A;
    font-weight: 700 !important;
}

/* ── 구분선 ── */
.divider {
    height: 2px;
    background: repeating-linear-gradient(90deg,
        #FFCC80 0,#FFCC80 8px,transparent 8px,transparent 14px);
    border: none;
    margin: 1.6rem 0;
}

/* ── 섹션 라벨 ── */
.section-label {
    font-family: 'Gaegu', cursive !important;
    font-size: 1.15rem !important;
    color: #5D3A1A !important;
    font-weight: 700 !important;
    margin: 0 0 0.7rem !important;
    display: flex;
    align-items: center;
    gap: 6px;
}

/* ── 폼 위젯 ── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stTextInput"] > div > div > input,
[data-testid="stTextArea"] > div > div > textarea {
    border-radius: 10px !important;
    border: 2px solid #E8C9A0 !important;
    background: #FFFDF7 !important;
    font-family: 'Nanum Gothic', sans-serif !important;
}
[data-testid="stSelectbox"] > div > div:focus-within,
[data-testid="stTextInput"] > div > div > input:focus,
[data-testid="stTextArea"] > div > div > textarea:focus {
    border-color: #FF8A65 !important;
    box-shadow: 0 0 0 3px #FF8A6520 !important;
}

/* ── DB 탐색 팝업 ── */
.db-panel {
    background: #FFFDF7;
    border-radius: 16px;
    border: 2px solid #E8C9A0;
    padding: 1.2rem 1.4rem;
    margin: 0.8rem 0 1rem;
}
.db-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 8px;
    margin-top: 0.8rem;
    max-height: 320px;
    overflow-y: auto;
}
.db-card {
    background: white;
    border: 1.5px solid #F0D9B8;
    border-radius: 10px;
    padding: 8px 10px;
    cursor: pointer;
    transition: border-color .15s, transform .15s;
    font-size: 0.8rem;
    line-height: 1.45;
}
.db-card:hover { border-color: #FF8A65; transform: translateY(-2px); }
.db-card .dc-title { font-weight: 800; color: #3D2B1F; }
.db-card .dc-author { color: #9E8070; font-size: 0.74rem; }
.db-card .dc-tags { margin-top: 4px; display:flex; flex-wrap:wrap; gap:3px; }
.dc-tag {
    background: #FFF3E0;
    color: #E65100;
    border-radius: 20px;
    padding: 1px 7px;
    font-size: 0.68rem;
    font-weight: 700;
    border: 1px solid #FFCC80;
}

/* ── 선택된 책 표시 ── */
.selected-book {
    background: linear-gradient(135deg, #FFF8E7, #FFF0F5);
    border: 2px solid #FFCC80;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    margin: 0.5rem 0 0;
    display: flex;
    align-items: flex-start;
    gap: 10px;
}
.sb-icon { font-size: 1.6rem; flex-shrink: 0; }
.sb-title { font-weight: 800; color: #3D2B1F; font-size: 0.9rem; }
.sb-meta  { color: #9E8070; font-size: 0.78rem; margin-top: 2px; }
.sb-tags  { margin-top: 5px; display:flex; flex-wrap:wrap; gap:3px; }
.sb-badge {
    background: #E8F5E9; color: #1B5E20;
    border: 1px solid #A5D6A7; border-radius: 20px;
    padding: 1px 8px; font-size: 0.68rem; font-weight: 800;
}

/* ── 생성 버튼 ── */
[data-testid="baseButton-primary"] {
    background: #FF7043 !important;
    border: none !important;
    border-radius: 50px !important;
    font-family: 'Gaegu', cursive !important;
    font-size: 1.1rem !important;
    font-weight: 700 !important;
    color: white !important;
    box-shadow: 0 4px 0 #BF360C !important;
    letter-spacing: 1px;
    transition: transform .1s, box-shadow .1s !important;
}
[data-testid="baseButton-primary"]:hover  { transform: translateY(-2px) !important; box-shadow: 0 6px 0 #BF360C !important; }
[data-testid="baseButton-primary"]:active { transform: translateY(2px)  !important; box-shadow: 0 2px 0 #BF360C !important; }
[data-testid="baseButton-secondary"] {
    border-radius: 50px !important;
    border: 2px solid #E8C9A0 !important;
    background: white !important;
    font-weight: 700 !important;
    transition: border-color .15s !important;
}
[data-testid="baseButton-secondary"]:hover { border-color: #FF8A65 !important; }

/* ── 결과 아코디언 ── */
[data-testid="stExpander"] {
    background: white !important;
    border: 2px solid #F0D9B8 !important;
    border-radius: 12px !important;
    margin-bottom: 6px !important;
    overflow: hidden;
}
[data-testid="stExpander"]:hover { border-color: #FF8A65 !important; }
[data-testid="stExpander"] summary {
    font-family: 'Gaegu', cursive !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    color: #3D2B1F !important;
    padding: 0.6rem 1rem !important;
}
[data-testid="stExpander"] > div > div { padding: 0 1rem 1rem !important; }

/* ── 다운로드 버튼 ── */
[data-testid="baseButton-download"] {
    border-radius: 50px !important;
    font-weight: 700 !important;
}

/* ── 탭 ── */
[data-testid="stTabs"] [role="tab"] {
    font-family: 'Gaegu', cursive !important;
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    border-radius: 10px 10px 0 0 !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #BF360C !important;
    border-bottom: 3px solid #FF7043 !important;
}

/* ── 반응형 ── */
@media (max-width: 600px) {
    .main .block-container { padding: 1rem 0.8rem 2rem !important; }
    .app-title { font-size: 1.5rem !important; }
    .db-grid { grid-template-columns: 1fr; }
}
</style>
"""

SYSTEM_PROMPT = """당신은 초등 초기 문해력, 그림책 수업, 질문 중심 수업 설계 전문가입니다.
초보 교사가 바로 사용할 수 있도록 구체적이고 실제적인 수업안을 작성합니다. 반드시 한국어로 작성합니다.
수업은 질문-대화-활동-성찰로 이어지게 설계합니다.
초기 문해력 요소(음운인식, 어휘, 이야기이해, 추론, 배경지식, 감정이해, 표현능력)를 반영합니다.
질문은 사실/추론/평가/삶연결 질문이 균형 있게 포함됩니다.
대화형 읽기(Dialogic Reading)의 PEER 절차를 반영합니다.
학생 발달 수준에 맞는 쉬운 언어를 사용합니다."""


def get_client():
    api_key = st.secrets.get("OPENAI_API_KEY", None) or os.getenv("OPENAI_API_KEY")
    if not api_key:
        st.error("OpenAI API 키가 설정되지 않았습니다.")
        st.stop()
    return OpenAI(api_key=api_key)


def build_prompt(grade, theme, book, lesson_time, student_context, book_info):
    book_ctx = ""
    if book_info:
        q = book_info.get("questions", {})
        book_ctx = f"""
[DB 그림책 정보]
작가: {book_info['author']} / 줄거리: {book_info['summary']}
문해력 요소: {', '.join(book_info['literacy_elements'])} / 활용 이유: {book_info['reason']}
참고 질문 — 사실: {' / '.join(q.get('사실',[])[:2])} | 추론: {' / '.join(q.get('추론',[])[:2])} | 삶연결: {q.get('삶연결',[''])[0]}
"""
    return f"""다음 조건으로 'AI 그림책 질문수업 설계안'을 작성해 주세요.

[조건]
학년: {grade} / 주제: {theme} / 그림책: {book} / 수업시간: {lesson_time}
학생특성: {student_context or '특별한 조건 없음'}
{book_ctx}

[출력 — 아래 9개 섹션을 반드시 포함]
## 1. 수업 개요
수업명 / 대상 학년 / 그림책 / 핵심 주제 / 초기 문해력 요소 / 수업 목표 3개

## 2. 그림책 활용 포인트
적합한 이유 / 글 요소 / 그림 요소 / 교사 주의사항

## 3. 질문 생성
읽기 전(3) / 읽는 중(5) / 읽은 후(5) — 사실/추론/평가/감정/작가의도/삶연결 유형 표시

## 4. 활동 생성
활동1 도입 / 활동2 중심 / 활동3 표현·정리 — 각각 목표/준비물/진행/발문/예상학생반응

## 5. 활동지 초안
제목 / 안내문장 / 문항 5개 / 표현 활동 1개

## 6. 지도안
도입/전개/정리 — 시간배분 / 교사발문 / 학생활동 / 유의점

## 7. 평가
관찰평가 기준 4개 / 자기평가 문항 3개 / 교사 피드백 예시 5개

## 8. 학부모 안내문
그림책·주제 소개 / 가정 대화 질문 3개 / 따뜻한 어조

## 9. AI 활용 팁
교사가 추가로 물어볼 수 있는 프롬프트 5개"""


def generate_lesson(prompt, model):
    client = get_client()
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": SYSTEM_PROMPT},
                  {"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=4096,
    )
    return resp.choices[0].message.content or ""


def make_docx(text, title):
    doc = Document()
    doc.styles["Normal"].font.name = "맑은 고딕"
    doc.styles["Normal"].font.size = Pt(10.5)
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    for line in text.splitlines():
        s = line.strip()
        if not s: doc.add_paragraph("")
        elif s.startswith("## "): doc.add_heading(s[3:], level=2)
        elif s.startswith("### "): doc.add_heading(s[4:], level=3)
        elif s.startswith("# "): doc.add_heading(s[2:], level=1)
        elif s.startswith(("- ","• ")):
            p = doc.add_paragraph(style="List Bullet"); p.add_run(s[2:])
        else: doc.add_paragraph(s)
    buf = BytesIO(); doc.save(buf); return buf.getvalue()


def parse_sections(text):
    """결과를 9개 섹션으로 분리"""
    section_titles = {
        "1": "📋 수업 개요",
        "2": "📖 그림책 활용 포인트",
        "3": "❓ 질문 생성",
        "4": "🎨 활동 생성",
        "5": "📝 활동지 초안",
        "6": "🗒️ 지도안",
        "7": "⭐ 평가",
        "8": "👨‍👩‍👧 학부모 안내문",
        "9": "🤖 AI 활용 팁",
    }
    import re
    parts = re.split(r'(?m)^##\s+(\d+)\.', text)
    sections = {}
    for i in range(1, len(parts), 2):
        num = parts[i].strip()
        content = parts[i+1].strip() if i+1 < len(parts) else ""
        # 첫 줄(제목)을 제거
        lines = content.split('\n')
        content = '\n'.join(lines[1:]).strip() if lines else content
        label = section_titles.get(num, f"섹션 {num}")
        sections[num] = (label, content)
    return sections


# ── 메인 ────────────────────────────────────────────────────────────
def main():
    st.markdown(CSS, unsafe_allow_html=True)

    # 헤더
    st.markdown("""
    <div class="app-header">
        <span class="app-icon">📚</span>
        <div class="app-title">AI 그림책 질문수업 설계기</div>
        <p class="app-sub">학년 · 주제 · 그림책을 고르면 수업 초안을 자동으로 만들어 드려요</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ── 입력 영역 ──────────────────────────────────────────────────
    st.markdown('<div class="section-label">🎯 수업 조건</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        grade = st.selectbox("학년", ["유치원","초등 1학년","초등 2학년","초등 3학년",
                                       "초등 4학년","초등 5학년","초등 6학년"], label_visibility="collapsed")
    with c2:
        theme = st.selectbox("수업 주제", db_all_themes(), label_visibility="collapsed")
    with c3:
        lesson_time = st.selectbox("수업 시간",
            ["40분","80분","120분","프로젝트 3차시","프로젝트 5차시"],
            label_visibility="collapsed")

    col_hint = st.columns([1,1,1])
    with col_hint[0]: st.caption("📌 학년")
    with col_hint[1]: st.caption("🎯 주제")
    with col_hint[2]: st.caption("⏰ 시간")

    student_context = st.text_area(
        "학생 특성 (선택)",
        placeholder="예: 1학년 입학 초기, 친구 관계가 서툰 편, 글쓰기 부담이 큼 등",
        height=72,
        label_visibility="collapsed",
    )
    st.caption("📝 학생 특성 (선택)")

    # ── 그림책 선택 ────────────────────────────────────────────────
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-label">📖 그림책 선택</div>', unsafe_allow_html=True)

    book_tab1, book_tab2 = st.tabs(["DB에서 찾기", "직접 입력"])
    book = ""
    book_info = None

    with book_tab1:
        # DB 탐색 토글
        show_db = st.toggle("🔍 DB 전체 탐색", value=False)

        if show_db:
            st.markdown('<div class="db-panel">', unsafe_allow_html=True)
            fc1, fc2 = st.columns(2)
            with fc1:
                ft = st.selectbox("주제", ["전체"] + db_all_themes(), key="db_theme")
            with fc2:
                fg = st.selectbox("학년", ["전체","유치원","초등 1학년","초등 2학년",
                                            "초등 3학년","초등 4학년"], key="db_grade")
            filtered = db_search(
                "" if ft == "전체" else ft,
                "" if fg == "전체" else fg,
            )
            st.caption(f"검색 결과 {len(filtered)}권")

            # 카드 그리드 (HTML)
            cards_html = '<div class="db-grid">'
            for b in filtered:
                tags = "".join(f'<span class="dc-tag">{t}</span>' for t in b["theme"][:2])
                cards_html += (
                    f'<div class="db-card">'
                    f'<div class="dc-title">{b["title"]}</div>'
                    f'<div class="dc-author">{b["author"]}</div>'
                    f'<div class="dc-tags">{tags}</div>'
                    f'</div>'
                )
            cards_html += "</div>"
            st.markdown(cards_html, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # 추천 드롭다운 (학년+주제 기준)
        recommended = db_search(theme=theme, grade=grade)
        if recommended:
            rec_titles = [b["title"] for b in recommended]
            selected = st.selectbox(
                f"추천 그림책 ({len(rec_titles)}권 — {grade} × {theme})",
                rec_titles,
            )
            book_info = db_get_by_title(selected)
            book = selected
            if book_info:
                tags_html = "".join(
                    f'<span class="sb-badge">{e}</span>'
                    for e in book_info["literacy_elements"]
                )
                st.markdown(
                    f'<div class="selected-book">'
                    f'<span class="sb-icon">📕</span>'
                    f'<div>'
                    f'<div class="sb-title">{book_info["title"]}</div>'
                    f'<div class="sb-meta">{book_info["author"]} · {book_info["summary"]}</div>'
                    f'<div class="sb-tags">{tags_html}</div>'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )
        else:
            st.info("선택 조건에 맞는 책이 없습니다. '직접 입력' 탭을 이용해 주세요.")

    with book_tab2:
        custom_book = st.text_input("그림책 제목", placeholder="예: 알사탕")
        if custom_book:
            book = custom_book
            book_info = db_get_by_title(custom_book)
            if book_info:
                st.success(f"✅ DB에 '{custom_book}'이 있어요! 질문 데이터를 활용합니다.")
            else:
                st.info("DB에 없는 책이에요. AI가 일반 지식으로 수업안을 만들게요.")

    # ── AI 모델 (숨김 — 고급 설정) ────────────────────────────────
    with st.expander("⚙️ 고급 설정", expanded=False):
        model = st.selectbox("AI 모델",
            ["gpt-4o-mini","gpt-4o","gpt-4.1-mini","gpt-4.1"], index=0)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 생성 버튼 ──────────────────────────────────────────────────
    generate_btn = st.button("✨ 수업 설계안 생성하기", type="primary", use_container_width=True)

    if generate_btn:
        if not book.strip():
            st.warning("⚠️ 그림책을 선택하거나 입력해 주세요.")
            st.stop()
        st.session_state.pop("result", None)
        st.session_state.pop("title", None)
        prompt = build_prompt(grade, theme, book, lesson_time, student_context, book_info)
        with st.spinner("🖍️ 수업 설계안을 만드는 중이에요..."):
            try:
                result = generate_lesson(prompt, model)
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
                st.stop()
        st.session_state["result"] = result
        st.session_state["title"] = f"{book}_{theme}_질문수업설계안"

    # ── 결과 출력 ──────────────────────────────────────────────────
    if "result" in st.session_state:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-label">✅ 생성 결과</div>', unsafe_allow_html=True)

        sections = parse_sections(st.session_state["result"])

        if sections:
            for num, (label, content) in sections.items():
                with st.expander(label, expanded=(num == "1")):
                    st.markdown(content)
        else:
            # 파싱 실패 시 전체 출력
            st.markdown(st.session_state["result"])

        st.markdown("<br>", unsafe_allow_html=True)
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📄 Word 다운로드",
                data=make_docx(st.session_state["result"], st.session_state["title"]),
                file_name=f"{st.session_state['title']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "📝 Markdown 다운로드",
                data=st.session_state["result"].encode("utf-8"),
                file_name=f"{st.session_state['title']}.md",
                mime="text/markdown",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
