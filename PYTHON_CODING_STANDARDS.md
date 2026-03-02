# 파이썬 코딩 표준 가이드

## 📋 목차
1. [프로젝트 구조](#프로젝트-구조)
2. [모듈 구성 원칙](#모듈-구성-원칙)
3. [명명 규칙](#명명-규칙)
4. [클래스 설계](#클래스-설계)
5. [함수 작성](#함수-작성)
6. [에러 처리](#에러-처리)
7. [문서화](#문서화)

---

## 프로젝트 구조

### 권장 디렉토리 구조
```
project/
├── core/                    # 핵심 비즈니스 로직
│   ├── report_check.py
│   ├── eco_check.py
│   └── eco_input.py
├── gui/                     # GUI 관련 코드
│   ├── common/             # 공통 GUI 컴포넌트
│   │   ├── __init__.py
│   │   ├── widgets.py      # 재사용 가능한 위젯
│   │   └── log_panel.py    # 로그 표시 패널
│   ├── report_check_gui.py
│   ├── eco_check_gui.py
│   └── eco_input_gui.py
├── utils/                   # 유틸리티 함수
│   ├── __init__.py
│   ├── worker_utils.py     # 백그라운드 작업 관련
│   └── file_utils.py       # 파일 처리 관련
└── main.py                  # 진입점
```

### 파일 분리 기준
- **100줄 이상**: 별도 파일로 분리 고려
- **공통 기능**: 재사용 가능한 유틸리티 모듈로 분리
- **UI와 로직**: 반드시 분리 (GUI 파일과 비즈니스 로직 파일)

---

## 모듈 구성 원칙

### 1. 단일 책임 원칙 (SRP)
각 모듈은 하나의 명확한 책임만 가져야 합니다.

**좋은 예:**
```python
# gui_common.py - GUI 공통 기능만
# worker_utils.py - 백그라운드 작업만
# file_utils.py - 파일 처리만
```

**나쁜 예:**
```python
# utils.py - GUI, 파일, 네트워크 등 모든 것이 섞임
```

### 2. 함수 그룹화
비슷한 기능의 함수들은 클래스로 묶거나 같은 모듈에 배치합니다.

**개선 전:**
```python
# 여기저기 흩어진 함수들
def clear_clipboard():
    pass

def find_excel():
    pass

def open_file():
    pass
```

**개선 후:**
```python
# file_utils.py
class ExcelFileHandler:
    """엑셀 파일 처리 관련 기능"""
    
    @staticmethod
    def find_excel(sample, folder):
        """엑셀 파일 찾기"""
        pass
    
    @staticmethod
    def clear_clipboard():
        """클립보드 초기화"""
        pass
    
    @staticmethod
    def open_workbook(filepath):
        """워크북 열기"""
        pass
```

### 3. 의존성 최소화
모듈 간 의존성을 최소화하여 결합도를 낮춥니다.

```python
# ✅ 좋은 예: 필요한 것만 import
from gui_common import LogPanel, ProgressButton

# ❌ 나쁜 예: 모든 것을 import
from gui_common import *
```

---

## 명명 규칙

### 변수와 함수
- **snake_case** 사용
- 의미를 명확히 전달하는 이름 사용

```python
# ✅ 좋은 예
user_name = "홍길동"
sample_list = ["S001", "S002"]

def calculate_total_price(items):
    pass

# ❌ 나쁜 예
un = "홍길동"
sl = ["S001", "S002"]

def calc(x):
    pass
```

### 클래스
- **PascalCase** 사용
- 명사형으로 작성

```python
# ✅ 좋은 예
class ReportCheckGUI:
    pass

class BackgroundWorker:
    pass

# ❌ 나쁜 예
class report_gui:
    pass

class ProcessData:  # 동사형은 피함
    pass
```

### 상수
- **UPPER_SNAKE_CASE** 사용

```python
# ✅ 좋은 예
MAX_RETRY_COUNT = 3
DEFAULT_TIMEOUT = 30
SRC_DIR = r"\\192.168.10.163\측정팀"

# ❌ 나쁜 예
maxRetry = 3
default_timeout = 30
```

### 비공개 멤버
- 앞에 언더스코어(`_`) 추가

```python
class MyClass:
    def __init__(self):
        self._internal_value = 0  # 비공개 속성
    
    def _helper_method(self):     # 비공개 메서드
        pass
    
    def public_method(self):      # 공개 메서드
        pass
```

---

## 클래스 설계

### 1. 하나의 클래스는 하나의 책임
각 클래스는 명확한 단일 책임을 가져야 합니다.

```python
class LogPanel:
    """콘솔 로그를 표시하는 패널 - 로그 표시만 담당"""
    
    def __init__(self, parent, height=12):
        self.parent = parent
        self.log_queue = queue.Queue()
        self._setup_ui()
    
    def start_pumping(self):
        """로그 큐에서 메시지 읽어 표시"""
        pass
    
    def restore_stdio(self):
        """stdout/stderr 원복"""
        pass
```

### 2. 초기화 메서드 구조화
복잡한 초기화는 여러 메서드로 분리합니다.

```python
class EcoInputGUI:
    def __init__(self):
        self.root = tk.Tk()
        self._setup_window()        # 윈도우 설정
        self._create_widgets()      # 위젯 생성
        self._setup_dependencies()  # 의존성 설정
    
    def _setup_window(self):
        """윈도우 기본 설정"""
        self.root.title("측정인 자동 입력")
        self.root.geometry("930x840")
    
    def _create_widgets(self):
        """모든 위젯 생성"""
        self._create_input_area()
        self._create_log_area()
    
    def _setup_dependencies(self):
        """위젯 간 의존성 설정"""
        self.tab2_var.trace_add("write", self._update_pdf_state)
```

### 3. 속성 그룹화
관련된 속성들은 함께 초기화합니다.

```python
class EcoInputGUI:
    def __init__(self):
        # UI 변수들
        self.tab1_var = tk.BooleanVar(value=True)
        self.tab2_var = tk.BooleanVar(value=True)
        self.pdf_var = tk.BooleanVar(value=True)
        
        # 위젯들
        self.entry_id = None
        self.entry_pw = None
        self.start_btn = None
```

---

## 함수 작성

### 1. 함수는 짧고 명확하게
- 한 함수는 한 가지 일만 수행
- 15~20줄을 넘지 않도록 노력

```python
# ✅ 좋은 예: 각 단계를 별도 함수로 분리
def process_samples(sample_list):
    """시료 처리 메인 함수"""
    validated_samples = validate_samples(sample_list)
    excel_files = find_excel_files(validated_samples)
    results = process_excel_files(excel_files)
    return results

def validate_samples(samples):
    """시료 검증"""
    return [s for s in samples if is_valid_sample(s)]

def find_excel_files(samples):
    """엑셀 파일 찾기"""
    pass

# ❌ 나쁜 예: 모든 것을 한 함수에
def process_samples(sample_list):
    """시료 처리"""
    # 검증
    validated = []
    for s in sample_list:
        if len(s) > 0 and s.isalnum():
            validated.append(s)
    
    # 파일 찾기
    files = []
    for s in validated:
        # 30줄의 파일 찾기 로직...
    
    # 파일 처리
    for f in files:
        # 40줄의 파일 처리 로직...
```

### 2. 매개변수는 적게
- 매개변수는 3개 이하 권장
- 많을 경우 dict나 클래스로 묶기

```python
# ✅ 좋은 예
def create_worker(root, task, callbacks):
    """
    Args:
        root: Tkinter root
        task: 실행할 작업
        callbacks: dict {"on_success": func, "on_error": func}
    """
    pass

# ❌ 나쁜 예
def create_worker(root, task, on_success, on_error, on_complete, use_com, timeout):
    pass
```

### 3. 기본값 활용
```python
def create_log_panel(parent, height=12, title="로그"):
    """
    로그 패널 생성
    
    Args:
        parent: 부모 위젯
        height: 높이 (기본: 12)
        title: 제목 (기본: "로그")
    """
    pass
```

### 4. 타입 힌트 사용 (Python 3.5+)
```python
from typing import List, Optional, Callable

def process_samples(
    samples: List[str],
    callback: Optional[Callable] = None
) -> List[str]:
    """
    시료 처리
    
    Args:
        samples: 시료 번호 리스트
        callback: 진행상황 콜백 함수 (선택)
    
    Returns:
        처리된 시료 리스트
    """
    pass
```

---

## 에러 처리

### 1. 구체적인 예외 처리
```python
# ✅ 좋은 예
try:
    with open(filepath, 'r') as f:
        data = f.read()
except FileNotFoundError:
    log("파일을 찾을 수 없습니다")
except PermissionError:
    log("파일 접근 권한이 없습니다")
except Exception as e:
    log(f"알 수 없는 오류: {e}")

# ❌ 나쁜 예
try:
    with open(filepath, 'r') as f:
        data = f.read()
except:  # 모든 예외를 무시
    pass
```

### 2. 리소스 정리
```python
# ✅ 좋은 예: with 문 사용
with open(filepath, 'r') as f:
    data = f.read()

# 또는 try-finally
try:
    f = open(filepath, 'r')
    data = f.read()
finally:
    f.close()

# ❌ 나쁜 예
f = open(filepath, 'r')
data = f.read()
# close() 누락
```

### 3. 백그라운드 작업의 에러 처리
```python
def worker():
    """백그라운드 작업"""
    try:
        # 작업 수행
        result = perform_task()
    except Exception as e:
        # 에러 로깅
        traceback.print_exc()
        # UI에 알림
        root.after(0, lambda: messagebox.showerror("오류", str(e)))
    finally:
        # 정리 작업
        cleanup_resources()
```

---

## 문서화

### 1. 모듈 문서화
파일 맨 위에 모듈 설명 추가

```python
# -*- coding: utf-8 -*-
"""
GUI 공통 유틸리티 모듈

이 모듈은 모든 GUI에서 사용하는 공통 기능을 제공합니다:
- LogPanel: 콘솔 로그 표시 패널
- QueueWriter: stdout/stderr 리다이렉팅
- ProgressButton: 상태 표시 버튼

사용 예:
    from gui_common import LogPanel
    
    log_panel = LogPanel(parent_frame)
    log_panel.start_pumping()
"""
```

### 2. 클래스 문서화
```python
class BackgroundWorker:
    """
    백그라운드에서 작업을 실행하는 워커 클래스
    
    이 클래스는 GUI를 블로킹하지 않고 무거운 작업을 
    별도 스레드에서 실행합니다.
    
    Attributes:
        root: Tkinter root 윈도우
        task: 실행할 작업 (callable)
        use_com: Excel COM 사용 여부
    
    Example:
        worker = BackgroundWorker(
            root,
            task=my_function,
            on_success=lambda: print("완료")
        )
        worker.start()
    """
    pass
```

### 3. 함수 문서화
```python
def find_excel(sample: str, folder: str) -> Optional[str]:
    """
    폴더에서 시료번호에 해당하는 엑셀 파일 찾기
    
    Args:
        sample: 시료번호 (예: "S001")
        folder: 검색할 폴더 경로
    
    Returns:
        찾은 파일의 전체 경로, 없으면 None
    
    Raises:
        OSError: 폴더 접근 실패 시
    
    Example:
        >>> find_excel("S001", r"C:\Data")
        'C:\\Data\\S001_report.xlsx'
    """
    pass
```

### 4. 주석 작성 원칙
- **왜(Why)**를 설명, **무엇(What)**은 코드로
- 복잡한 로직은 단계별로 주석 추가
- TODO, FIXME, NOTE 등 태그 활용

```python
# ✅ 좋은 예: 왜 이렇게 하는지 설명
# Excel COM은 스레드마다 초기화 필요
if pythoncom is not None:
    pythoncom.CoInitialize()

# NOTE: 탭2가 꺼지면 PDF도 자동으로 비활성화
if not tab2_var.get():
    pdf_var.set(False)

# FIXME: 대용량 파일에서 성능 저하 발생
for row in range(1, 10000):
    process_row(row)

# ❌ 나쁜 예: 코드를 그대로 반복
# i를 1부터 10까지 반복
for i in range(1, 11):
    print(i)  # i를 출력
```

---

## 실전 적용 체크리스트

### 새 모듈 생성 시
- [ ] 모듈 문서화 추가
- [ ] import 정리 (표준 라이브러리 → 서드파티 → 로컬)
- [ ] 상수를 파일 상단에 정의

### 새 클래스 생성 시
- [ ] 클래스 문서화 추가
- [ ] `__init__` 메서드 구조화 (여러 메서드로 분리)
- [ ] 비공개 멤버에 `_` 접두사 추가

### 새 함수 생성 시
- [ ] 함수 문서화 추가
- [ ] 한 가지 일만 수행하는지 확인
- [ ] 매개변수 3개 이하 유지
- [ ] 타입 힌트 추가

### 리팩토링 시
- [ ] 중복 코드 제거
- [ ] 공통 기능 유틸리티로 추출
- [ ] 함수/클래스 분리
- [ ] 테스트 코드 작성

---

## 참고 자료
- [PEP 8 – Style Guide for Python Code](https://peps.python.org/pep-0008/)
- [Google Python Style Guide](https://google.github.io/styleguide/pyguide.html)
- [Clean Code in Python](https://www.oreilly.com/library/view/clean-code-in/9781800560215/)
