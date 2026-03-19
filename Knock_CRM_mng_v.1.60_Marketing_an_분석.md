# Knock CRM Management v.1.60 - Marketing Analysis VBA 분석 문서

## 1. 파일 개요

### 기본 정보
- **파일명**: Knock_CRM_mng_v.1.60_Marketing_an.xlsm
- **버전**: v.1.60
- **파일 타입**: OpenXML (Excel Macro-Enabled Workbook)
- **주요 목적**: 마케팅 데이터(광고비, 매출) 관리 및 DB 연동

### 프로젝트 설명
이 VBA 프로젝트는 Knock CRM 시스템의 마케팅 분석을 위한 Excel 기반 데이터 관리 도구입니다. Oracle 데이터베이스와 연동하여 광고비 현황과 차트번호별 매출 데이터를 조회, 삽입, 삭제하는 기능을 제공합니다.

---

## 2. 전체 구조

### 모듈 구성
| 모듈/클래스 | 타입 | 설명 | 코드 존재 여부 |
|------------|------|------|---------------|
| 현재_통합_문서.cls | Workbook Class | 통합 문서 클래스 | Empty |
| Sheet1.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet2.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet3.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet4.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet5.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet6.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet7.cls | Worksheet Class | 워크시트 클래스 | Empty |
| Sheet8.cls | Worksheet Class | 워크시트 클래스 | 코드 존재 (Option Explicit만) |
| Sheet9.cls | Worksheet Class | 워크시트 클래스 | Empty |
| **Module1.bas** | Standard Module | **메인 로직 모듈** | **코드 존재** |

### 아키텍처
- **데이터베이스**: Oracle (ODBC DSN 연결)
- **연결 방식**: ADODB (ActiveX Data Objects)
- **주요 워크시트**:
  - `업체별 광고비 현황`: 광고비 데이터 관리
  - `차트번호별 매출현황`: 매출 데이터 관리
  - `작업`: 임시 작업 및 데이터 조회
  - `SQL`: SQL 쿼리 저장소

---

## 3. 전체 Sub/Function 목록

| 함수명 | 매개변수 | 주요 역할 | 호출 빈도 |
|--------|----------|-----------|----------|
| `AddCostInsert()` | 없음 | 업체별 광고비 데이터를 Excel에서 DB로 일괄 삽입 | Medium |
| `SalesDataInsert()` | 없음 | 차트번호별 매출 데이터를 Excel에서 DB로 일괄 삽입 | Medium |
| `RunQueryAndLoadResult()` | 없음 | 광고비 현황 조회 쿼리 실행 및 결과를 Excel에 로드 (파라미터: H2, H3) | High |
| `RunQuery_Load_Work()` | 없음 | 작업 시트로 데이터 조회 (파라미터: S2, S3, 결과: A~I열) | High |
| `RunQuery_ChartSalesByChartNo()` | 없음 | 차트번호별 매출현황 조회 (파라미터: T2, T3, 결과: B~K열) | High |
| `Run_Delete_ChartSalesStatus()` | 없음 | CHART_SALES_STATUS 테이블 전체 데이터 삭제 (초기화) | Low |
| `Run_Delete_AdCostCompany()` | 없음 | AD_COST_COMPANY 테이블 전체 데이터 삭제 (초기화) | Low |

---

## 4. 주요 상수 및 전역 변수

### 데이터베이스 연결 정보 (하드코딩)
```vba
DSN: knock_crm_real
UID: knock_crm
PWD: kkptcmr!@34
```

### 타임아웃 설정
| 항목 | 값 | 설명 |
|------|-----|------|
| ConnectionTimeout | 300초 (5분) | DB 연결 타임아웃 |
| CommandTimeout | 300초 (5분) | SQL 명령 실행 타임아웃 |

### 전역 변수
이 파일에는 전역 변수가 선언되어 있지 않습니다. 모든 변수는 프로시저 내에서 로컬로 선언됩니다.

**주석 처리된 전역 변수**:
```vba
'Private Connection As ADODB.Connection  (라인 32 - 사용 안 함)
```

---

## 5. DB 스키마 정보

### 5.1. AD_COST_COMPANY 테이블 (광고비 업체별 현황)
| 컬럼명 | 타입 | 설명 | 소스 |
|--------|------|------|------|
| DB_INPUT_DATE | DATE/VARCHAR | DB 입력일자 | Cells(i, 1) |
| DB_SRC_1 | VARCHAR | DB 소스 1 (광고 채널 등) | Cells(i, 2) |
| DB_SRC_2 | VARCHAR | DB 소스 2 (세부 채널) | Cells(i, 3) |
| EVENT | VARCHAR | 이벤트 구분 | Cells(i, 4) |
| AD_COST | NUMBER | 광고비 금액 | Cells(i, 5) |

**관련 프로시저**:
- `AddCostInsert()`: INSERT 작업
- `RunQueryAndLoadResult()`: SELECT 조회
- `Run_Delete_AdCostCompany()`: DELETE 작업 (SQL!C5)

### 5.2. CHART_SALES_STATUS 테이블 (차트번호별 매출현황)
| 컬럼명 | 타입 | 설명 | 소스 |
|--------|------|------|------|
| INPUT_DATE | VARCHAR(8) | 입력일자 (YYYYMMDD) | Cells(i, 1) - Format으로 변환 |
| RESERV_DATE | VARCHAR | 예약일자 | Cells(i, 2) |
| RESERV_TIME | VARCHAR | 예약시간 | Cells(i, 3) |
| CHART_NO | VARCHAR | 차트번호 (고유 식별자) | Cells(i, 4) |
| CLIENT_NAME | VARCHAR | 고객명 | Cells(i, 5) |
| PHONE_NO | VARCHAR | 전화번호 | Cells(i, 6) |
| EVENT_TYPE | VARCHAR | 이벤트 유형 | Cells(i, 7) |
| DB_SRC_2 | VARCHAR | DB 소스2 | Cells(i, 8) |
| REF_DATE | VARCHAR | 참조일자 | Cells(i, 9) |
| DB_INPUT_DATE | VARCHAR | DB 입력일자 | Cells(i, 10) |
| TM_NAME | VARCHAR | TM 이름 (텔레마케터) | Cells(i, 11) |
| CONSENT_SALES | NUMBER | 동의 매출액 | Cells(i, 17) |

**관련 프로시저**:
- `SalesDataInsert()`: INSERT 작업
- `RunQuery_ChartSalesByChartNo()`: SELECT 조회 (SQL!C4)
- `Run_Delete_ChartSalesStatus()`: DELETE 작업 (SQL!C6)

### 5.3. SQL 쿼리 저장 위치 (SQL 시트)
| 셀 위치 | 쿼리 용도 | 사용 프로시저 |
|---------|----------|--------------|
| C2 | 광고비 현황 조회 (파라미터: :param01, :param02) | RunQueryAndLoadResult() |
| C3 | 작업 시트 데이터 조회 (파라미터: :param01, :param02) | RunQuery_Load_Work() |
| C4 | 차트번호별 매출현황 조회 (파라미터: :param01, :param02) | RunQuery_ChartSalesByChartNo() |
| C5 | AD_COST_COMPANY 삭제 쿼리 | Run_Delete_AdCostCompany() |
| C6 | CHART_SALES_STATUS 삭제 쿼리 | Run_Delete_ChartSalesStatus() |

---

## 6. 핵심 비즈니스 로직 설명

### 6.1. 데이터 삽입 프로세스

#### AddCostInsert() - 광고비 일괄 등록
```
[프로세스 흐름]
1. "업체별 광고비 현황" 시트 활성화
2. DB 연결 (ADODB.Connection, 타임아웃 300초)
3. 마지막 행 감지 (A열 기준)
4. 2행부터 마지막 행까지 반복:
   - A열~E열 데이터 읽기
   - INSERT 쿼리 동적 생성
   - DB 실행 (Connection.Execute)
   - Debug.Print로 로그 출력
5. COMMIT 실행
6. 연결 종료 및 완료 메시지
```

**특징**:
- 날짜 포맷 변환 없이 직접 삽입
- 에러 핸들링 없음 (실패 시 중단)

#### SalesDataInsert() - 매출 데이터 일괄 등록
```
[프로세스 흐름]
1. "차트번호별 매출현황" 시트 활성화
2. DB 연결
3. 마지막 행 감지
4. 2행부터 반복:
   - A열(입력일자)를 YYYYMMDD 포맷으로 변환
   - 12개 컬럼 데이터 매핑 (주석: 일부 날짜 포맷 변환 코드는 비활성화)
   - INSERT 쿼리 실행
5. COMMIT 및 종료
```

**특징**:
- `Format(Cells(i, 1).Value, "yyyymmdd")` 날짜 변환 적용
- 컬럼 건너뛰기 있음 (12~16열 제외, 17열 사용)

### 6.2. 쿼리 조회 및 결과 로드 프로세스

#### 공통 패턴
모든 조회 프로시저는 다음 패턴을 따릅니다:
```
1. 파라미터 시트에서 날짜 범위 읽기 (param01, param02)
2. SQL 시트에서 쿼리 템플릿 읽기
3. 바인드 변수 치환 (:param01, :param02 → 실제 값)
4. 세미콜론 제거 (Oracle 호환)
5. ADODB.Recordset으로 쿼리 실행
6. 컬럼명 + 데이터를 지정 범위에 출력
7. AutoFit으로 열 너비 자동 조정
```

#### RunQueryAndLoadResult() - 광고비 조회
- **파라미터 위치**: H2 (시작일), H3 (종료일)
- **SQL 위치**: SQL!C2
- **결과 출력**: A1:E열 (컬럼명 1행, 데이터 2행~)
- **클리어 범위**: A1:E1000000

#### RunQuery_Load_Work() - 작업 시트 조회
- **파라미터 위치**: S2, S3 (작업 시트)
- **SQL 위치**: SQL!C3
- **결과 출력**: A2:I열 (컬럼명 2행, 데이터 3행~)
- **최대 컬럼**: 9개 제한
- **CursorLocation**: adUseClient(3) 사용

#### RunQuery_ChartSalesByChartNo() - 매출현황 조회
- **파라미터 위치**: T2, T3 (차트번호별 매출현황 시트)
- **SQL 위치**: SQL!C4
- **결과 출력**: B1:K열 (B열부터 시작, 10개 컬럼)
- **특징**: A열은 사용하지 않음 (다른 용도로 예약된 듯)

### 6.3. 데이터 삭제 프로세스

#### 공통 안전장치
```vba
userResponse = MsgBox("테이블을 초기화 합니다." & vbCrLf & _
                      "계속 진행하시겠습니까?", vbYesNo + vbQuestion, "삭제 확인")
If userResponse = vbNo Then Exit Sub
```

#### Run_Delete_ChartSalesStatus()
- SQL!C6 셀에서 DELETE 쿼리 읽기
- 사용자 확인 후 실행
- COMMIT 자동 실행

#### Run_Delete_AdCostCompany()
- SQL!C5 셀에서 DELETE 쿼리 읽기
- 동일한 확인 프로세스

---

## 7. 다른 파일과의 차별점

### 7.1. 마케팅 분석 특화
이 파일은 **마케팅 분석(Marketing_an)**에 특화된 VBA 모듈로, 다음 기능에 집중합니다:
- 광고비 현황 관리 (AD_COST_COMPANY)
- 매출 현황 추적 (CHART_SALES_STATUS)
- 광고 투자 대비 성과 분석을 위한 데이터 수집

### 7.2. 예상되는 Knock CRM 시스템의 다른 파일들
| 예상 파일 | 차별점 |
|-----------|--------|
| **Knock_CRM_mng_v.1.60_Sales.vba** | 영업 관리 중심 (계약, 고객 관리) |
| **Knock_CRM_mng_v.1.60_TM.vba** | 텔레마케팅 관리 (콜 로그, 상담 내역) |
| **Knock_CRM_mng_v.1.60_Admin.vba** | 관리자 기능 (권한, 시스템 설정) |

### 7.3. 이 파일의 고유 특징

#### 1. 광고 채널 추적
- `DB_SRC_1`, `DB_SRC_2` 컬럼으로 광고 출처 이중 분류
- 이벤트별 광고비 분리 관리 (`EVENT` 컬럼)

#### 2. 날짜 파라미터 기반 조회
모든 조회 함수가 2개의 날짜 파라미터를 받아 기간별 분석 지원:
```vba
sqlText = Replace(sqlText, ":param01", "'" & param01 & "'")
sqlText = Replace(sqlText, ":param02", "'" & param02 & "'")
```

#### 3. 시트별 특화된 출력 범위
| 시트 | 컬럼 범위 | 시작 행 | 용도 |
|------|----------|---------|------|
| 업체별 광고비 현황 | A~E | 1행 | 광고비 집계 |
| 작업 | A~I | 2행 | 임시 분석 작업 |
| 차트번호별 매출현황 | B~K | 1행 | 매출 상세 (A열 제외) |

#### 4. 데이터 무결성 관리
- 삭제 작업 시 사용자 확인 메시지 (2단계 확인)
- COMMIT 명시적 실행 (트랜잭션 관리)
- 파라미터 검증 (빈 값 체크)

#### 5. SQL 외부화 전략
쿼리를 VBA 코드에 하드코딩하지 않고 "SQL" 시트의 셀에 저장:
- **장점**: 쿼리 수정 시 코드 변경 불필요, 비개발자도 수정 가능
- **단점**: 셀 위치 의존성 (C2~C6 고정)

---

## 8. 추가 분석 및 개선 제안

### 8.1. 보안 이슈
**심각**: 데이터베이스 인증 정보 평문 노출
```vba
conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
```
**권장 조치**:
- Windows 자격 증명 관리자 사용
- 환경 변수 또는 암호화된 설정 파일 활용

### 8.2. 에러 핸들링 부재
모든 프로시저에 `On Error` 구문이 없어 다음 위험 존재:
- DB 연결 실패 시 프로그램 중단
- 데이터 타입 불일치 시 런타임 에러
- 트랜잭션 롤백 불가

**권장 조치**:
```vba
On Error GoTo ErrorHandler
' ... 코드 ...
Exit Sub

ErrorHandler:
    If Not conn Is Nothing Then
        conn.Execute "ROLLBACK"
        conn.Close
    End If
    MsgBox "오류: " & Err.Description, vbCritical
End Sub
```

### 8.3. 성능 개선 기회

#### 현재 방식 (행 단위 INSERT)
```vba
For i = 2 To lastRow
    sql = "INSERT INTO ..."
    Connection.Execute sql  ' 매번 DB 통신
Next i
```

#### 개선안 (일괄 INSERT)
- Bulk INSERT 또는 Prepared Statement 사용
- 100~1000행 단위 배치 처리
- Array를 사용한 메모리 버퍼링

### 8.4. 코드 중복 제거
DB 연결 코드가 모든 프로시저에 반복됨
```vba
Set conn = CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 300
conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
```

**권장**: 공통 함수 분리
```vba
Function GetConnection() As Object
    Set GetConnection = CreateObject("ADODB.Connection")
    With GetConnection
        .ConnectionTimeout = 300
        .CommandTimeout = 300
        .Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    End With
End Function
```

### 8.5. 날짜 포맷 불일치
- `AddCostInsert()`: 날짜 변환 없음
- `SalesDataInsert()`: `Format(*, "yyyymmdd")` 사용

일관성 있는 날짜 처리 필요.

---

## 9. 사용 시나리오

### 일반적인 워크플로우
```
[월간 마케팅 분석 프로세스]

1. 광고비 데이터 수집
   ├─ Excel "업체별 광고비 현황" 시트에 데이터 입력
   └─ AddCostInsert() 실행 → DB 저장

2. 매출 데이터 수집
   ├─ Excel "차트번호별 매출현황" 시트에 데이터 입력
   └─ SalesDataInsert() 실행 → DB 저장

3. 기간별 분석
   ├─ H2, H3에 조회 기간 입력 (예: 20250801, 20250831)
   ├─ RunQueryAndLoadResult() 실행 → 광고비 현황 조회
   ├─ T2, T3에 날짜 입력
   └─ RunQuery_ChartSalesByChartNo() 실행 → 매출 현황 조회

4. 데이터 정리 (필요 시)
   ├─ Run_Delete_AdCostCompany() → 광고비 테이블 초기화
   └─ Run_Delete_ChartSalesStatus() → 매출 테이블 초기화
```

---

## 10. 의존성 및 요구사항

### 시스템 요구사항
- **Excel 버전**: 2010 이상 (OpenXML 지원)
- **VBA 참조**: Microsoft ActiveX Data Objects 2.x Library
- **ODBC 드라이버**: Oracle ODBC Driver 설치 필요
- **DSN 설정**: knock_crm_real DSN 구성 필요

### 네트워크 요구사항
- Oracle DB 서버 접근 권한
- 포트: 일반적으로 1521 (TNS Listener)

---

## 11. 결론

이 VBA 파일은 Knock CRM 시스템의 마케팅 분석 모듈로서, 광고비와 매출 데이터를 효과적으로 관리하는 도구입니다. Excel의 직관적인 인터페이스와 Oracle DB의 강력한 데이터 관리 능력을 결합하여, 마케팅 담당자가 기술적 지식 없이도 데이터를 입력하고 조회할 수 있도록 설계되었습니다.

**강점**:
- 간단한 UI (Excel 시트 기반)
- SQL 외부화로 유지보수 용이
- 명확한 데이터 흐름 (Insert/Select/Delete)

**개선 필요 영역**:
- 보안 강화 (인증 정보 보호)
- 에러 핸들링 추가
- 성능 최적화 (일괄 처리)
- 로깅 기능 추가

전체적으로 실용적이고 목적에 부합하는 코드이지만, 프로덕션 환경에서는 위에서 제안한 개선 사항들을 적용하는 것이 권장됩니다.
