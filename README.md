# 엑셀 변경점 추적기

두 개의 엑셀 파일을 넣으면 git diff처럼 무엇이 추가·삭제·수정됐는지 보여줍니다.
파일은 브라우저 밖으로 나가지 않습니다.

## 설치

파이썬 3.9 이상만 있으면 됩니다. 파일 7개를 한 폴더에 두고 그 폴더에서 실행하세요.

```bash
cd excel-diff
python serve.py
```

브라우저가 `http://127.0.0.1:8000` 으로 열립니다. 파이썬 런타임(Pyodide)과 openpyxl은
CDN/PyPI가 아니라 이 저장소의 `pyodide/`, `wheels/` 폴더에 이미 같이 들어 있어서,
바로 로컬 파일에서 읽어옵니다 — 인터넷이 없어도, 처음 여는 사람이어도 몇 초 안에 뜹니다.

`serve.py`는 표준 라이브러리만 쓰기 때문에 **브라우저에서만 쓸 거라면 설치할 패키지가 없습니다.**

터미널 CLI(`python diff_engine.py ...`)도 쓰려면 openpyxl이 필요합니다.

```bash
python -m venv .venv
source .venv/bin/activate        # 윈도우: .venv\Scripts\activate
pip install -r requirements.txt
```

포트 8000이 이미 쓰이고 있으면 `python serve.py --port 9000` 처럼 바꾸면 됩니다.
`python3`로 실행해야 하는 환경이라면 명령어의 `python`을 `python3`로 바꿔주세요.

## 왜 파일이 서버로 안 가나

`serve.py`는 화면 파일(html, css, js, 그리고 파이썬 엔진 소스)만 내려주는 정적 서버입니다.
업로드 경로가 아예 없고, POST 요청은 405로 막아놨습니다.

엑셀을 올리면 브라우저는 이렇게 처리합니다.

```
파일 선택 → File.arrayBuffer()로 바이트를 읽음
          → Web Worker로 넘김  (같은 컴퓨터, 같은 브라우저 안)
          → Pyodide(WebAssembly로 컴파일된 CPython)가 받음
          → openpyxl로 파싱하고 diff_engine.py가 비교
          → 결과 JSON만 화면으로 돌아옴
```

파이썬이 프론트엔드(브라우저 안)에서도, 백엔드(비교 로직)에서도 돌아갑니다.
`diff_engine.py`는 브라우저와 일반 CPython에서 똑같이 동작하는 파일 하나입니다.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `diff_engine.py` | 비교 엔진. 브라우저와 터미널에서 공용으로 씁니다 |
| `worker.js` | Pyodide를 백그라운드 스레드에 띄우고 엔진을 호출합니다 |
| `app.js` | 파일 선택, 결과 렌더링, 내려받기 |
| `index.html`, `styles.css` | 화면 |
| `serve.py` | 정적 파일 서버 (표준 라이브러리만 씀) |
| `requirements.txt` | CLI로 쓸 때 필요한 패키지 |
| `pyodide/` | 브라우저에서 파이썬을 돌리는 런타임(Pyodide) 본체. CDN 대신 여기서 읽습니다 |
| `wheels/` | openpyxl / et_xmlfile wheel. PyPI 대신 여기서 설치합니다 |

## 무엇을 잡아내나

- **시트**: 추가, 삭제
- **열**: 추가, 삭제, 위치 이동(이전 `C`열 → 이후 `B`열처럼 표시)
- **행**: 추가, 삭제, 수정
- **셀**: 값 변경 (이전 값을 취소선으로, 이후 값을 아래에 표시)

행과 열은 위치가 아니라 **내용**으로 맞춥니다. 중간에 행을 하나 끼워 넣어도
그 아래 행들이 전부 "바뀐 행"이 되지 않고, 끼워 넣은 한 줄만 추가로 잡힙니다.

동작 방식은 이렇습니다.

1. 각 열의 값 전체를 하나의 키로 만들어 `difflib`으로 열을 정렬합니다.
2. 추가·삭제된 열을 뺀 나머지 열의 값만으로 행 키를 만들어 다시 행을 정렬합니다.
3. 교체된 구간은 행끼리 얼마나 닮았는지 계산해서, 닮았으면 "수정된 행",
   전혀 다르면 "삭제 + 추가"로 나눕니다.

## 옵션

| 옵션 | 설명 |
|---|---|
| 비교 기준 | **계산된 값**(기본) 또는 **수식**. `=SUM(A1:A9)`가 그대로인데 결과만 달라진 경우를 구분할 때 바꿉니다 |
| 보기 | 자동 / 변경된 부분만(±3행) / 전체 행. 400행이 넘으면 자동으로 변경된 부분만 보여줍니다 |
| 대소문자 무시 | `Apple`과 `APPLE`을 같게 봅니다 |
| 공백 무시 | 셀 안의 모든 공백을 지우고 비교합니다 |

## 터미널에서 쓰기

브라우저에 올리기 버거운 큰 파일은 CLI가 빠릅니다.

```bash
python diff_engine.py 이전.xlsx 이후.xlsx
python diff_engine.py 이전.xlsx 이후.xlsx --report 리포트.xlsx --json 결과.json
python diff_engine.py 이전.xlsx 이후.xlsx --formula --ignore-space
```

`report_xlsx()`가 만드는 리포트는 요약 / 변경 목록 / 행·열 변경 세 시트로 되어 있고,
화면의 **엑셀 리포트 저장** 버튼도 같은 함수를 씁니다.

## 잘 안 될 때

화면이 특정 단계에서 멈추면 브라우저 개발자 도구(F12)의 콘솔을 보세요.
단계마다 `[엑셀 diff]` 로그가 찍히고, 실패하면 그 자리에서 이유가 보입니다.
각 단계에 제한 시간이 걸려 있어서 끝없이 멈춰 있지는 않습니다.

| 멈춘 단계 | 원인 | 해결 |
|---|---|---|
| 파이썬 런타임 내려받는 중 | `pyodide/` 폴더가 없거나 파일이 빠짐 | 아래 "번들 다시 받기"로 `pyodide/`를 채우기 |
| 엑셀 라이브러리 설치 중 | `wheels/` 폴더가 없거나 파일이 빠짐 | 아래 "번들 다시 받기"로 `wheels/`를 채우기 |
| 비교 엔진 올리는 중 | `diff_engine.py`가 같은 폴더에 없음 | 주소창에 `http://127.0.0.1:8000/diff_engine.py`를 쳐서 소스가 보이는지 확인 |

## 정적 호스팅에 배포하기

이 저장소는 이미 완전히 서버리스입니다. `serve.py`는 로컬 개발용 정적 파일 서버일 뿐,
비교는 전부 브라우저 안(Web Worker + Pyodide)에서 일어나고 파일은 어디로도 전송되지 않습니다.
그래서 폴더 전체(특히 `index.html`, `app.js`, `worker.js`, `styles.css`, `diff_engine.py`,
`pyodide/`, `wheels/`)를 그대로 정적 호스팅에 올리면 됩니다 — 백엔드도, 빌드 과정도 필요 없습니다.

- **GitHub Pages**: 이 저장소를 그대로 Pages로 켜면 됩니다
- **Netlify / Vercel / Cloudflare Pages**: 빌드 명령 없이 "정적 파일 배포"로 이 폴더를 그대로 업로드

`pyodide/`, `wheels/`를 함께 올려두면 방문자가 CDN이나 PyPI를 거치지 않고
같은 도메인에서 바로 런타임을 받으므로, 첫 방문도 빠르고 사내망/폐쇄망에서도 그대로 동작합니다.

## 번들 다시 받기 (버전을 올리거나 다시 받아야 할 때)

**엑셀 라이브러리**(openpyxl):

```bash
pip download openpyxl -d wheels --no-deps
pip download et_xmlfile -d wheels --no-deps
```

**파이썬 런타임**(Pyodide, jsDelivr 기준 `worker.js`의 `PYODIDE_VERSION`과 맞추기):

```bash
V=0.26.4
mkdir -p pyodide
for f in pyodide.js pyodide.mjs pyodide.asm.js pyodide.asm.wasm pyodide-lock.json \
         python_stdlib.zip micropip-0.6.0-py3-none-any.whl packaging-23.2-py3-none-any.whl; do
  curl -sfo "pyodide/$f" "https://cdn.jsdelivr.net/pyodide/v$V/full/$f"
done
```

micropip/packaging 파일명(버전)은 Pyodide 버전마다 다를 수 있습니다 —
`pyodide-lock.json`에서 `"micropip"`, `"packaging"` 항목의 `file_name`을 확인하고 맞춰 받으세요.

두 경우 모두 코드는 손대지 않아도 됩니다. `worker.js`는 `pyodide/`, `wheels/`에서 먼저 찾고
그 폴더에 없을 때만 CDN/PyPI로 넘어갑니다.

`index.html`의 Google Fonts 링크도 지우면 완전히 오프라인으로 돕니다.

## 한계

- `.xlsx`, `.xlsm`, `.xltx`만 읽습니다. 예전 `.xls`는 엑셀에서 한 번 저장해 변환하세요
- 값과 수식만 비교합니다. 글꼴, 색, 셀 병합, 조건부 서식, 차트, 이미지 변경은 잡지 않습니다
- 값 기준 비교는 엑셀이 저장해 둔 계산 결과를 읽습니다. 엑셀로 열지 않고 프로그램이
  만든 파일은 계산 결과가 비어 있을 수 있는데, 이때는 수식 텍스트로 대신 비교합니다
- 한 시트당 3만 행, 2천 열까지 브라우저에서 처리합니다. 그 이상은 CLI를 쓰세요
- 시트는 이름으로 짝지어집니다. 시트 이름을 바꾸면 삭제 + 추가로 보입니다
