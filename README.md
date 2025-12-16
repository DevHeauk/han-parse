# Han-Parse

한글 파일(HWP) 파싱 및 표 처리 도구

## 설치

```bash
pip install -r requirements.txt
```

## 주요 기능

- ✅ 한글 파일에서 텍스트 추출
- ✅ **표 파싱 및 구조 추출**
- ✅ **표 데이터를 JSON/CSV로 저장**
- ✅ **JSON/CSV 데이터로 표 재구성**
- ✅ 파일 메타데이터 읽기
- ✅ OLE 파일 구조 분석

## 사용법

### 1. 기본 파싱

```bash
# 텍스트와 표 모두 파싱
python han_parser.py example.hwp

# 표만 파싱
python han_parser.py example.hwp --tables
```

### 2. 표 데이터 저장

```bash
# 표를 JSON으로 저장
python han_parser.py example.hwp --save-json tables.json

# 표를 CSV로 저장 (각 표가 개별 파일로)
python han_parser.py example.hwp --save-csv tables_output/
```

### 3. 표 데이터 재구성

```bash
# JSON에서 표 데이터를 읽어서 한글 파일 생성
python table_reconstructor.py create-json tables.json output.hwp [template.hwp]

# CSV에서 표 데이터를 읽어서 한글 파일 생성
python table_reconstructor.py create-csv tables_output/ output.hwp [template.hwp]
```

### 4. 표 데이터 수정

```bash
# JSON 파일의 표 데이터 수정
# 형식: <표인덱스> <행> <열> <새값>
python table_reconstructor.py edit tables.json 0 1 2 "새로운 값"
```

### 5. 표 데이터 병합

```bash
# 두 개의 JSON 파일 병합
python table_reconstructor.py merge tables1.json tables2.json merged.json
```

## Python 코드에서 사용

### 기본 파싱

```python
from han_parser import parse_hwp, parse_hwp_simple, parse_tables

# 텍스트 추출
text = parse_hwp("example.hwp")
print(text)

# 파일 정보 추출
info = parse_hwp_simple("example.hwp")
print(info)

# 표 파싱
tables = parse_tables("example.hwp")
for idx, table in enumerate(tables):
    print(f"표 {idx + 1}: {table['row_count']}행 x {table['col_count']}열")
    for row in table['rows']:
        print(row)
```

### 표 데이터 저장/로드

```python
from han_parser import (
    parse_tables, 
    save_tables_to_json, 
    save_tables_to_csv,
    load_tables_from_json,
    load_tables_from_csv
)

# 표 파싱 및 저장
tables = parse_tables("example.hwp")
save_tables_to_json(tables, "tables.json")
save_tables_to_csv(tables, "tables_csv/")

# 표 데이터 로드
tables = load_tables_from_json("tables.json")
tables = load_tables_from_csv("tables_csv/")
```

### 표 데이터 수정 및 재구성

```python
from table_reconstructor import (
    edit_table_data,
    create_hwp_from_tables_json,
    merge_tables
)

# 표 데이터 수정
edit_table_data("tables.json", table_index=0, row=1, col=2, new_value="새 값")

# 표 데이터로 한글 파일 생성
create_hwp_from_tables_json("tables.json", "output.hwp", template_path="template.hwp")

# 표 데이터 병합
merge_tables("tables1.json", "tables2.json", "merged.json")
```

## 표 데이터 형식

표 데이터는 다음과 같은 JSON 구조를 가집니다:

```json
[
  {
    "rows": [
      ["헤더1", "헤더2", "헤더3"],
      ["데이터1", "데이터2", "데이터3"],
      ["데이터4", "데이터5", "데이터6"]
    ],
    "row_count": 3,
    "col_count": 3,
    "section": 0,
    "paragraph": 5
  }
]
```

## 주의사항

1. **표 파싱 제한**: 한글 파일의 복잡한 구조로 인해 일부 표는 완벽하게 파싱되지 않을 수 있습니다.

2. **재구성 제한**: `pyhwp`는 주로 읽기 전용이므로, 완전한 한글 파일 재구성은 제한적일 수 있습니다. 
   - 표 데이터는 JSON/CSV로 추출하여 수정한 후
   - 한글 프로그램에서 수동으로 삽입하거나
   - 한글과컴퓨터 API를 사용하는 것을 권장합니다.

3. **템플릿 사용**: 재구성 시 템플릿 한글 파일을 사용하면 더 나은 결과를 얻을 수 있습니다.

## 의존성

- `pyhwp>=0.1b12`: 한글 파일 파싱 라이브러리
- `olefile>=0.46`: OLE 파일 형식 처리
- Python 표준 라이브러리: `json`, `csv` (별도 설치 불필요)

## 예제 워크플로우

1. **표 추출 및 저장**
   ```bash
   python han_parser.py document.hwp --save-json tables.json
   ```

2. **표 데이터 수정** (Python 코드 또는 JSON 직접 편집)
   ```python
   from table_reconstructor import edit_table_data
   edit_table_data("tables.json", 0, 1, 0, "수정된 값")
   ```

3. **수정된 표로 재구성** (또는 수동으로 한글에 삽입)
   ```bash
   python table_reconstructor.py create-json tables.json output.hwp template.hwp
   ```

## 웹 인터페이스 사용법

### 웹 서버 실행

```bash
python app.py
```

서버가 실행되면 브라우저에서 `http://localhost:5000`으로 접속하세요.

### 웹 기능

1. **파일 업로드**
   - 드래그 앤 드롭 또는 클릭하여 한글 파일 업로드
   - 자동으로 텍스트와 표 파싱

2. **표 편집**
   - 웹 인터페이스에서 표 셀을 직접 편집
   - 여러 표가 있는 경우 드롭다운으로 선택
   - 실시간 편집 가능

3. **파일 다운로드**
   - 수정된 표 데이터로 한글 파일 생성 및 다운로드
   - JSON 형식으로도 다운로드 가능

### 웹 API 엔드포인트

- `POST /api/upload` - 한글 파일 업로드 및 파싱
- `GET /api/tables/<session_id>` - 표 데이터 가져오기
- `POST /api/tables/<session_id>` - 표 데이터 업데이트
- `POST /api/download/<session_id>` - 한글 파일 다운로드
- `GET /api/download-json/<session_id>` - JSON 파일 다운로드

## 라이선스

MIT

