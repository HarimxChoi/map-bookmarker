# map-bookmarker

**엑셀/CSV 주소 목록을 카카오맵, 네이버지도 즐겨찾기에 자동으로 등록해주는 프로그램**

주소 하나하나 검색해서 즐겨찾기 누르는 반복 작업, 이제 안 해도 됩니다.
엑셀 파일 하나만 준비하면 수백 개 주소를 자동으로 즐겨찾기에 등록합니다.

---

## 이런 분들에게 유용합니다

- 배달/퀵서비스 라이더 (담당 구역 고객 주소 등록)
- 방문 영업직 (고객사, 거래처 주소 관리)
- 부동산 임장 (매물 주소 일괄 저장)
- 설문/방문 조사원 (조사 대상 주소 목록)
- 학습지/방문 교사 (담당 학생 가정 주소)
- 방문 요양/의료 (환자 방문 주소 관리)

---

## 주요 기능

| 기능 | 설명 |
|------|------|
| 파일 지원 | Excel (.xlsx, .xls), CSV |
| 지도 플랫폼 | 카카오맵, 네이버지도 (동시 또는 개별 실행 가능) |
| GUI | 파일 선택하면 컬럼 자동 감지, 드롭다운으로 선택 |
| 폴더 관리 | 즐겨찾기 폴더 자동 생성/기존 폴더 선택 |
| 필터 | 특정 조건으로 등록할 행만 골라내기 |
| 중단/재시작 | 도중에 끊겨도 이미 등록된 건 자동으로 건너뜀 |
| 2차 인증 | 카카오톡 인증, 네이버 캡차 등 자동 대기 |
| 주소 정제 | 아파트 동/호수, 괄호 내용 자동 정리 |

---

## 설치 및 실행

### 방법 1: EXE 파일 (설치 없이 바로 실행)

> Python을 모르는 분도 사용 가능

1. [Google Drive에서 다운로드](https://drive.google.com/file/d/19zS21-VOu6PDV0_n5fAjP1C_1maSL14S/view?usp=sharing)
2. 압축 풀기
3. `MapFavoriteRegistrar.exe` 더블클릭
4. 첫 실행 시 브라우저(Chromium) 자동 설치 (1~2분)

### 방법 2: 소스코드 (개발자용)

**1단계: 다운로드**
```
git clone https://github.com/HarimxChoi/map-bookmarker
cd map-bookmarker
```

**2단계: 설치**

Windows:
```
install.bat
```

또는 직접:
```
pip install -r requirements.txt
python -m playwright install chromium
```

**3단계: 실행**

GUI 실행:
```
python run_gui.py
```

CLI 실행:
```
python src/main.py                  # 전체 실행
python src/main.py --kakao-only     # 카카오맵만
python src/main.py --naver-only     # 네이버지도만
python src/main.py --dry-run        # 미리보기 (등록 안 함)
python src/main.py --limit 5        # 5개만 테스트
```

---

## 사용 방법

### 1. 엑셀/CSV 파일 준비

아래처럼 **주소가 포함된 파일**이면 됩니다. 컬럼명은 아무거나 상관없습니다.

| 이름 | 주소 | 상태 | 메모 |
|------|------|------|------|
| 홍길동 | 서울 강남구 테헤란로 101 | 미방문 | 주차 가능 |
| 이순신 | 경기 성남시 분당구 내정로 186 | 미방문 | 1층 |
| 김유신 | 서울 종로구 사직로 161 | 완료 | |

### 2. GUI에서 설정

1. 프로그램 실행
2. **파일 선택** - 엑셀/CSV 파일 불러오기
3. **주소 컬럼 선택** - 드롭다운에서 주소가 있는 컬럼 선택
4. **즐겨찾기명 컬럼 선택** - (선택) 안 하면 주소가 이름으로 사용됨
5. **플랫폼 탭** - 카카오/네이버 계정 정보 입력, 폴더명 설정
6. **Start** 클릭

### 3. 로그인 인증

프로그램이 브라우저를 열고 자동으로 로그인합니다.

- **카카오맵**: 카카오톡 알림이 오면 핸드폰에서 승인 (최대 5분 대기)
- **네이버지도**: 캡차가 나오면 브라우저에서 직접 입력 (최대 5분 대기)

인증 완료 후 자동으로 등록이 시작됩니다.

### 4. 완료

등록이 끝나면 카카오맵/네이버지도 앱에서 즐겨찾기를 확인하세요.

---

## 필터 기능

특정 조건으로 등록할 행만 골라낼 수 있습니다.

GUI에서 **필터 추가** 버튼으로 설정하거나, config.yaml에서 직접 작성:

```yaml
filters:
  # "상태" 컬럼에 "완료"나 "취소"가 있으면 제외
  - column: "상태"
    not_contains: ["완료", "취소"]

  # "지역" 컬럼에 "분당구"나 "수지구"가 포함된 행만
  - column: "지역"
    contains: ["분당구", "수지구"]

  # "금액" 컬럼이 10000 이상인 행만
  - column: "금액"
    min: 10000
```

---

## 설정 파일 (config.yaml)

GUI를 사용하면 자동으로 설정되지만, 직접 수정도 가능합니다.

```yaml
# 입력 파일
input:
  file: "data/주소목록.xlsx"
  sheet: "Sheet1"
  header_row: 1

# 컬럼 매핑
columns:
  name: "이름"
  address: "주소"

# 즐겨찾기 이름 포맷 ({컬럼명}으로 조합 가능)
bookmark_name: "{이름}"

# 카카오맵
kakao:
  enabled: true
  id: "카카오_이메일"
  password: "비밀번호"
  folder: "내 폴더명"

# 네이버지도
naver:
  enabled: true
  id: "네이버_아이디"
  password: "비밀번호"
  folder: "내 폴더명"

# 실행 옵션
options:
  headless: false    # false: 브라우저 화면 보임 (권장)
  delay_ms: 800      # 등록 간격 (ms). 너무 빠르면 차단 위험
  max_retry: 3       # 실패 시 재시도 횟수
  resume: true       # 중단 후 재시작 시 이미 등록된 건 건너뜀
```

---

## 프로젝트 구조

```
map-bookmarker/
├── src/
│   ├── main.py              # 핵심 로직 (카카오맵/네이버지도 자동화)
│   ├── gui.py               # GUI (tkinter)
│   └── browser_connector.py # 브라우저 연결
├── config/
│   └── config.yaml          # 설정 파일
├── data/                    # 입력 파일 폴더
├── logs/                    # 실행 로그
├── run_gui.py               # GUI 실행 진입점
├── run.bat                  # Windows 실행 메뉴
├── install.bat              # Windows 원클릭 설치
└── requirements.txt
```

---

## 주의사항

- **등록 속도**: `delay_ms`를 300 이하로 설정하면 계정이 일시 차단될 수 있습니다. 800ms 이상을 권장합니다.
- **UI 변경**: 카카오맵/네이버지도가 UI를 업데이트하면 셀렉터가 변경되어 동작하지 않을 수 있습니다. Issues에 제보해주세요.
- **개인정보**: config.yaml에 계정 정보가 포함되므로 **절대 공개 저장소에 올리지 마세요**. `.gitignore`에 포함되어 있습니다.

---

## 기술 스택

- Python 3.10+
- Playwright (브라우저 자동화)
- tkinter (GUI)
- pandas, openpyxl (엑셀 처리)
- PyInstaller (EXE 패키징)

---

## License

MIT - 자유롭게 사용, 수정, 배포 가능합니다.
