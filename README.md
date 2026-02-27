# Excel Sequential Data Copier (엑셀 순차 복사기)

대량의 엑셀 데이터를 **단축키(Shift) 하나로 순서대로 복사**해주는 자동화 도구입니다.
사내 보안 환경(VBA 객체 제한 등)에서도 문제없이 작동하도록 설계되었습니다.

---

## 주요 기능

| 단축키 | 동작 |
|--------|------|
| `왼쪽 Shift` | 다음 데이터를 클립보드에 복사 (`Ctrl+V`로 바로 붙여넣기) |
| `오른쪽 Shift` | 이전 데이터로 되돌아가기 |
| `ESC` | 매크로 종료 |

- **자동 반올림**: 소수점 자릿수를 열별로 다르게 설정 가능
- **다중 특별 열**: 소수점 자릿수가 다른 열을 여러 개 지정 가능
- **보안 환경 대응**: `DataObject` 없이 Windows API로 직접 클립보드 제어

---

## 설치 방법

1. 엑셀에서 `Alt + F11` → VBA 편집기 열기
2. 상단 메뉴 `삽입` → `모듈` 클릭
3. [code.txt](code.txt)의 내용 전체를 복사하여 붙여넣기
4. `Alt + F8` → `StartSequentialCopy` 선택 → **[실행]**
5. 준비 완료 메시지가 뜨면 시작

---

## 설정 방법

`code.txt` 상단의 **사용자 설정 구역**만 수정하면 됩니다.
코드 편집기(`Alt + F11`)를 열면 파일 맨 위에 바로 보입니다.

```vba
' ⚡ [사용자 설정 구역] 이 부분만 수정하세요! ⚡

Private Const START_ROW_NUM  As Long   = 1    ' 데이터 시작 행
Private Const START_COL_CHAR As String = "B"  ' 복사 시작 열
Private Const END_COL_CHAR   As String = "G"  ' 복사 끝 열

Private Const DIGITS_DEFAULT  As Integer = 3  ' 기본 소수점 자릿수
Private Const DIGITS_SPECIAL  As Integer = 5  ' 특별 열 소수점 자릿수

Private Const SPECIAL_COLS As String = "E"    ' 특별 자릿수 열 (쉼표로 여러 개 가능)
```

### 설정 항목 설명

| 항목 | 기본값 | 설명 |
|------|--------|------|
| `START_ROW_NUM` | `1` | 데이터가 시작하는 행 번호. 첫 행이 제목(헤더)이면 `2`로 변경 |
| `START_COL_CHAR` | `"B"` | 복사를 시작할 열 (알파벳) |
| `END_COL_CHAR` | `"G"` | 복사를 끝낼 열 (알파벳) |
| `DIGITS_DEFAULT` | `3` | 기본 소수점 자릿수. `0`으로 설정하면 정수로 반올림 |
| `DIGITS_SPECIAL` | `5` | 특별 열에 적용할 소수점 자릿수 |
| `SPECIAL_COLS` | `"E"` | 특별 자릿수를 적용할 열. 여러 열은 `"E,F"` 형식으로 작성. 없애려면 `""` |

---

## 설정 예시

### 예시 1 — 기본 사용 (A열~F열, 소수점 2자리)

```vba
Private Const START_ROW_NUM  As Long   = 2    ' 1행이 제목이므로 2부터 시작
Private Const START_COL_CHAR As String = "A"
Private Const END_COL_CHAR   As String = "F"
Private Const DIGITS_DEFAULT  As Integer = 2
Private Const DIGITS_SPECIAL  As Integer = 5
Private Const SPECIAL_COLS As String = ""     ' 특별 열 없음
```

### 예시 2 — 특별 열이 여러 개인 경우 (C, E열을 소수점 6자리로)

```vba
Private Const START_ROW_NUM  As Long   = 1
Private Const START_COL_CHAR As String = "B"
Private Const END_COL_CHAR   As String = "H"
Private Const DIGITS_DEFAULT  As Integer = 3
Private Const DIGITS_SPECIAL  As Integer = 6
Private Const SPECIAL_COLS As String = "C,E"  ' C열과 E열에 6자리 적용
```

### 예시 3 — 숫자만 있고 소수점이 필요 없는 경우

```vba
Private Const START_ROW_NUM  As Long   = 1
Private Const START_COL_CHAR As String = "A"
Private Const END_COL_CHAR   As String = "D"
Private Const DIGITS_DEFAULT  As Integer = 0  ' 정수로 반올림
Private Const DIGITS_SPECIAL  As Integer = 0
Private Const SPECIAL_COLS As String = ""
```

---

## 자주 묻는 질문

**Q. 작동이 안 돼요 / 보안 오류가 납니다.**
A. VBA 편집기에서 `도구` → `매크로 보안` → 보안 수준을 **낮음** 또는 **중간**으로 설정하세요.

**Q. 첫 번째 행(제목)이 복사됩니다.**
A. `START_ROW_NUM`을 `2`로 변경하세요.

**Q. 소수점이 안 잘려요 / 원래 값 그대로 나와요.**
A. 해당 셀이 텍스트 형식으로 저장된 것입니다. 셀 서식을 `숫자`로 바꾼 뒤 값을 다시 입력하세요.

**Q. macOS에서 쓸 수 있나요?**
A. 불가능합니다. Windows 전용 API를 사용합니다.
