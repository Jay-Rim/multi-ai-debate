# AI 토론

Perplexity, Gemini, ChatGPT, Claude의 **웹 서비스 화면을 Playwright로 자동 조작**하여 근거를 수집하고 토론하게 만드는 Streamlit 애플리케이션입니다.

별도의 AI API를 호출하지 않으므로 API 사용료는 발생하지 않습니다. 단, 각 AI 서비스의 웹사이트에 로그인할 수 있는 본인 계정이 필요합니다.

## 토론 진행 방식

진행 순서는 다음과 같이 고정되어 있습니다.

1. **Perplexity - 검색 비서/근거 수집**
   - 최신 공식 자료, 기사, 제품 정보, 논문과 원문 링크를 수집합니다.
   - 공식 출처와 사용자 리뷰·커뮤니티 의견을 분리해 팩트팩을 만듭니다.
   - 최종 판단에는 참여하지 않습니다.

2. **Gemini - 팩트 교차검증 + 근거 기반 발산**
   - Perplexity의 핵심 사실을 `검증됨 / 상충함 / 미검증`으로 분류합니다.
   - 가능하면 다른 독립 출처를 찾아 교차검증하고 누락된 사실과 반대 근거를 추가합니다.
   - 새로운 사실·수치에는 원문 URL을 붙이며, 출처가 없으면 가설 또는 추가 확인 필요로 표시합니다.
   - 최종 결론은 내리지 않고 가능한 선택지와 숨은 변수를 넓게 탐색합니다.

3. **ChatGPT - 검증**
   - Gemini 주장의 논리적 허점과 근거 부족을 검토합니다.
   - Perplexity 출처의 신뢰도, 비용, 시간, 운영 위험과 실행 가능성을 따집니다.
   - Perplexity와 Gemini가 같은 원출처를 반복한 경우 독립 교차검증으로 인정하지 않습니다.

4. **Claude - 수렴**
   - Gemini와 ChatGPT의 의견을 종합합니다.
   - 최종 판단과 실행 방안을 제시합니다.

5. **ChatGPT - 최종 반대자**
   - Claude의 결론이 틀렸을 가능성을 다시 검토합니다.
   - 성급한 결론, 누락된 위험, 과대평가된 실행 가능성을 확인합니다.

6. **Claude - 수정 최종안**
   - ChatGPT의 최종 반론을 반영해 최종 보고서를 수정합니다.

완료된 결과는 Word 문서로 로컬 PC에 저장됩니다. 설정한 경우 Telegram으로 완료 알림도 받을 수 있습니다.

## 주요 기능

- Perplexity, Gemini, ChatGPT, Claude 웹사이트 자동 제어
- Perplexity를 이용한 최신 정보·공식 출처 팩트팩 생성
- 토론 세트 수 설정
- 토론 도중 같은 채팅 입력창으로 진행자 의견 개입
- 각 AI 답변이 끝날 때까지 DOM 변화를 감지하여 대기
- 로그인 또는 CAPTCHA가 필요할 때 수동 해결 가능
- Claude 수정 최종안을 Word 문서로 저장
- 체크 해제 시 토론 기록과 결과 문서 로컬 보관
- 기본값으로 질문·답변을 PC에 저장하지 않는 로컬 비공개 모드
- ngrok을 이용한 스마트폰·외부 네트워크 접속
- Telegram 완료 알림
- ChatGPT 작업 대화 링크 표시

## 알아둘 점

이 프로젝트는 공식 API가 아니라 AI 서비스의 웹 화면을 자동화합니다.

- AI 사이트의 화면 구조가 바뀌면 셀렉터가 작동하지 않을 수 있습니다.
- 각 AI 서비스의 이용약관과 정책을 확인하고 본인 책임하에 사용하세요.
- 로그인, 계정 확인, CAPTCHA는 Chrome 창에서 직접 처리해야 할 수 있습니다.
- 자동화 브라우저가 화면에 표시되므로 데스크톱 로그인 세션이 필요합니다.
- Windows 환경에서 가장 많이 테스트했습니다.
- 웹 구독 상태와 로그인 세션은 사용자가 직접 관리해야 합니다.

## 실행 환경

- Python 3.11 이상 권장
- Google Chrome 또는 Microsoft Edge
- Windows 10/11 권장
- Perplexity, Gemini, ChatGPT, Claude에 접속 가능한 계정
- 선택 사항: 외부 접속을 위한 ngrok 계정
- 선택 사항: 완료 알림을 위한 Telegram 봇

## 설치 방법

### 1. 저장소 복제

PowerShell을 열고 다음 명령을 실행합니다.

```powershell
git clone https://github.com/Jay-Rim/multi-ai-debate.git
cd multi-ai-debate
```

### 2. 가상환경 만들기

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

PowerShell 실행 정책 때문에 활성화가 거부되면 현재 창에서 다음 명령을 먼저 실행합니다.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### 3. 패키지 설치

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

기본 설정은 PC에 설치된 Chrome을 사용합니다. Playwright 전용 Chromium을 사용하려면 `.env`의 `BROWSER_CHANNEL`을 비우고 다음 명령도 실행합니다.

```powershell
python -m playwright install chromium
```

### 4. 환경설정 파일 만들기

```powershell
Copy-Item .env.example .env
```

메모장이나 코드 편집기로 `.env` 파일을 엽니다.

```powershell
notepad .env
```

## `.env` 설정

최소 실행에는 대부분 기본값을 사용할 수 있습니다.

```dotenv
# ngrok 외부 접속을 사용할 때만 입력
NGROK_AUTHTOKEN=

# 자동화 전용 Chrome 프로필 경로
CHROME_USER_DATA_DIR=C:/Users/사용자이름/ai_debate_profile

# chrome, msedge 또는 빈 값
BROWSER_CHANNEL=chrome

STREAMLIT_PORT=8501

# 선택 사항: Google 계정 선택 화면에서 자동으로 선택할 이메일
GOOGLE_ACCOUNT=

# 선택 사항: Telegram 완료 알림
TELEGRAM_BOT_TOKEN=
TELEGRAM_CHAT_ID=
```

### Chrome 프로필 설정 주의사항

평소 사용하는 Chrome 프로필 폴더를 직접 지정하지 마세요. 이미 실행 중인 Chrome과 프로필 잠금이 충돌할 수 있습니다.

다음처럼 **AI 토론 전용 새 폴더**를 지정하는 것이 안전합니다.

```dotenv
CHROME_USER_DATA_DIR=C:/Users/사용자이름/ai_debate_profile
```

이 폴더는 첫 실행 시 자동으로 생성됩니다.

## 실행 방법

```powershell
streamlit run app.py
```

정상적으로 실행되면 터미널에 다음 주소가 표시됩니다.

```text
http://localhost:8501
```

브라우저에서 이 주소를 열면 AI 토론 화면이 나타납니다.

동시에 자동화용 Chrome 창에 Perplexity, Gemini, ChatGPT, Claude 탭이 열립니다. 첫 실행에서는 각 탭에 직접 로그인하고 Perplexity의 Cloudflare 보안 확인이 나타나면 수동으로 통과하세요. 로그인 상태는 전용 Chrome 프로필 폴더에 저장되어 다음 실행부터 재사용됩니다.

## 사용 방법

1. Streamlit 화면 하단에서 기록 설정과 토론 세트 수를 선택합니다.
2. 기본 체크된 `이 PC에 질문·답변 기록 남기지 않기`를 유지하면 로컬 비공개 모드로 실행됩니다.
3. 채팅 입력창에 토론 주제를 입력합니다.
4. Perplexity가 최신 정보와 공식 출처를 검색해 팩트팩을 만듭니다.
5. Gemini가 팩트팩을 교차검증하고 누락 팩트·상충 자료·대안을 추가합니다.
6. ChatGPT가 출처의 독립성, 주장과 원문의 일치 여부, 논리와 현실성을 검증합니다.
7. Claude가 1차 결론을 작성합니다.
8. ChatGPT가 최종 반대자 관점에서 결론을 다시 검토합니다.
9. Claude가 수정 최종안을 작성합니다.
10. Word 문서가 생성되고 결과 화면에 다운로드 버튼이 표시됩니다.

토론 진행 중 채팅 입력창에 추가 의견을 쓰면 다음 AI 발화 전에 **Human-in-the-loop 진행자 개입**으로 전달됩니다.

## 로컬 기록 설정

`이 앱의 질문·답변 기록을 PC에 남기지 않기` 체크박스는 기본적으로 켜져 있습니다.

체크한 상태에서는 다음 항목을 디스크에 저장하지 않습니다.

- 질문과 AI 답변 히스토리 JSON
- `logs/debate_events.log`의 질문·답변 실행 로그
- `out/` 폴더의 Word 결과 파일

결과는 실행 중인 앱 메모리에만 유지되며 화면에서 Word 파일을 직접 다운로드할 수 있습니다. 앱이나 서버를 재시작하면 메모리에 있던 결과는 사라집니다.

체크를 풀면 기존 방식으로 `history/`, `logs/`, `out/`에 기록을 저장합니다.

이 설정은 **이 애플리케이션이 `history/`, `logs/`, `out/`에 남기는 기록만 제어**합니다. 지속 로그인을 위한 자동화 Chrome 프로필에는 방문 기록이나 캐시가 남을 수 있습니다. Perplexity, Gemini, ChatGPT, Claude 서버에 대화가 저장되는지도 각 서비스의 계정·대화 기록·임시 채팅 설정에 따라 별도로 결정됩니다.

## 토론 세트 수

토론 세트 수가 `1`이면 다음 순서로 진행됩니다.

```text
Perplexity 팩트팩
→ Gemini → ChatGPT → Claude 초안
→ ChatGPT 최종 반대자 검토
→ Claude 수정 최종안
→ Word 생성
```

토론 세트 수가 `2`이면 Gemini, ChatGPT, Claude의 토론을 두 번 반복한 뒤 최종 검토 단계로 넘어갑니다.

```text
Perplexity 팩트팩
→ Gemini → ChatGPT → Claude
→ Gemini → ChatGPT → Claude 초안
→ ChatGPT 최종 반대자 검토
→ Claude 수정 최종안
→ Word 생성
```

## 외부 접속 설정

스마트폰이나 다른 네트워크에서 접속하려면 ngrok을 사용할 수 있습니다.

### 1. ngrok 토큰 발급

[ngrok 대시보드](https://dashboard.ngrok.com/get-started/your-authtoken)에서 인증 토큰을 발급받아 `.env`에 입력합니다.

```dotenv
NGROK_AUTHTOKEN=본인의_ngrok_토큰
```

### 2. 터널 실행

새 PowerShell 창에서 프로젝트 폴더로 이동한 뒤 실행합니다.

```powershell
.\.venv\Scripts\Activate.ps1
python start_tunnel.py
```

터미널에 다음과 같은 외부 URL이 표시됩니다.

```text
NGROK_URL=https://example.ngrok-free.app
```

이 주소를 스마트폰 브라우저에서 열면 됩니다. PC의 Streamlit과 자동화 Chrome은 계속 실행 중이어야 합니다.

외부 URL은 로컬의 `ngrok_url.txt`에 저장되지만 Git에는 올라가지 않습니다.

## Telegram 완료 알림

Telegram 봇을 사용하면 토론이 끝났을 때 스마트폰으로 완료 알림을 받을 수 있습니다.

`.env`에 다음 값을 설정합니다.

```dotenv
TELEGRAM_BOT_TOKEN=봇_토큰
TELEGRAM_CHAT_ID=채팅_ID
```

알림에는 다음 정보만 포함됩니다.

- 토론 완료 여부
- 토론 주제
- 저장된 파일명
- 확인 가능한 경우 ChatGPT 대화 링크

토론 본문과 Word 파일은 Telegram으로 전송하지 않습니다.

## 생성되는 로컬 데이터

실행 중 다음 폴더와 파일이 생성됩니다.

| 경로 | 용도 |
|---|---|
| `history/` | 완료된 토론 기록 |
| `logs/` | 실행 및 자동화 로그 |
| `out/` | 생성된 Word 결과물 |
| `ngrok_url.txt` | 현재 ngrok 주소 |
| `ai_debate_profile/` | Chrome 로그인 세션 |

이 항목들은 `.gitignore`에 포함되어 GitHub에 올라가지 않습니다.

특히 Chrome 프로필에는 로그인 세션이 들어 있으므로 다른 사람과 공유하면 안 됩니다.

## 자주 발생하는 문제

### Playwright 브라우저 실행 파일이 없다는 오류

예시:

```text
Executable doesn't exist
Please run: playwright install
```

해결 방법:

```powershell
python -m playwright install chromium
```

또는 `.env`에서 설치된 Chrome을 사용합니다.

```dotenv
BROWSER_CHANNEL=chrome
```

### 같은 Chrome 프로필을 사용 중이라는 오류

자동화 전용 Chrome 창을 모두 닫고 다시 실행하세요. 일반 Chrome 프로필이 아니라 별도 프로필 경로를 사용해야 합니다.

### 로그인 또는 CAPTCHA에서 멈춤

자동화 Chrome 창을 확인하고 직접 로그인하거나 CAPTCHA를 해결하세요. 특히 Perplexity는 최초 접속 시 Cloudflare 보안 확인 화면이 나타날 수 있습니다. 입력창이 나타나면 자동화가 다시 진행됩니다.

### 답변이 중간에 잘림

AI 사이트의 응답 DOM 구조나 복사 버튼이 변경되었을 수 있습니다. `logs/debate_events.log`에서 마지막으로 처리된 AI와 오류 메시지를 확인하세요.

### 외부 주소가 열리지 않음

1. Streamlit이 `8501` 포트에서 실행 중인지 확인합니다.
2. `start_tunnel.py` 프로세스가 실행 중인지 확인합니다.
3. ngrok을 종료한 뒤 다시 실행해 새 터널을 만듭니다.

로컬 상태 확인:

```powershell
Invoke-WebRequest http://127.0.0.1:8501/_stcore/health
```

정상이라면 `ok`가 반환됩니다.

## 보안 주의사항

- `.env` 파일을 Git에 커밋하지 마세요.
- Telegram 봇 토큰과 ngrok 토큰을 공개하지 마세요.
- Chrome 프로필 폴더를 업로드하지 마세요.
- 생성된 토론 문서에 사내 정보나 개인정보가 포함될 수 있으므로 공개 저장소에 올리지 마세요.
- 외부 ngrok URL을 공유하면 해당 주소를 아는 사람이 Streamlit 화면에 접근할 수 있습니다.

## 프로젝트 구조

```text
multi-ai-debate/
├─ app.py                 # Streamlit UI와 Playwright 토론 자동화
├─ start_tunnel.py        # ngrok 터널 실행
├─ requirements.txt       # Python 패키지 목록
├─ .env.example           # 환경설정 예시
├─ .gitignore             # 민감·실행 데이터 제외
└─ .streamlit/
   └─ config.toml         # Streamlit 화면 설정
```

## 라이선스

MIT License
