# Knock CRM Marketing 분석 문서

## 1. 파일 개요

### 기본 정보
- **파일명**: Knock_CRM_mng_v.1.60_Marketing.xlsm
- **버전**: v.1.60
- **목적**: 치과 CRM 시스템의 마케팅 DB 관리 및 배포를 위한 Excel VBA 매크로 애플리케이션
- **총 라인 수**: 10,717 lines
- **대상 사용자**: 마케터

### 주요 기능
- 마케팅 DB 입력, 조회, 배포 관리
- TM(텔레마케터) 콜 로그 관리
- DentWeb(치과 예약 시스템) 연동
- DB 결산 및 통계 분석
- 광고비 및 매출 현황 관리

## 2. 전체 구조

### 2.1 주요 모듈 구성

| 모듈명 | 타입 | 주요 역할 |
|--------|------|-----------|
| ThisWorkbook.cls | Workbook Class | 사용자 로그인, 초기 설정 |
| DB_Agent.cls | Class Module | ADODB 연결, 트랜잭션 관리 |
| Sheet_조회.bas | Sheet Module | 각종 데이터 조회 기능 |
| Sheet_DB_배포.bas | Sheet Module | TM에게 DB 배포 및 관리 |
| Sheet_InputDB_작업.bas | Sheet Module | 다양한 소스의 DB 표준화 |
| Sheet_InputDB_입력.bas | Sheet Module | 표준화된 DB를 시스템에 입력 |
| Sheet_DentWeb_연동.bas | Sheet Module | 덴트웹 예약 시스템 데이터 동기화 |
| Sheet_DB_결산.bas | Sheet Module | DB 사용 현황 결산 |
| Sheet_통계.bas | Sheet Module | TM별, DB업체별 성과 분석 |
| Sheet_구DB_작업.bas | Sheet Module | 과거 DB 재활용 |
| Util.bas | Standard Module | 날짜 변환, 문자열 처리 등 |
| SQL_Wrapper.bas | Standard Module | SQL 파라미터 바인딩 |

### 2.2 데이터베이스 연결 정보

```vba
DSN=knock_crm_real
uid=knock_crm
pwd=kkptcmr!@34
```

**보조 DB 연결** (DentWeb):
```vba
DSN=dentweb
uid=sa
pwd=Q3xzJiwpv2zC
```

## 3. Sub/Function 목록

### 3.1 워크북 초기화 및 로그인

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Workbook_open | Private Sub | 워크북 열릴 때 자동 실행 |
| Workbook_Initialize | Sub | 사용자 로그인 처리 및 권한 확인 |

### 3.2 조회 기능 (Sheet_조회.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_조회_Query | Public Sub | 조회 타입별 분기 처리 |
| Sheet_조회_DB_회수_내역_조회 | Public Sub | DB 회수 내역 조회 |
| Sheet_조회_Call_Log_조회 | Public Sub | TM 콜 로그 조회 |
| Sheet_조회_연락처_history_전체_조회 | Public Sub | 특정 전화번호의 전체 이력 조회 |
| Sheet_조회_DB_배포_내역_조회 | Public Sub | DB 배포 내역 조회 |
| Sheet_조회_INPUT_DB_source별_입력시간_조회 | Public Sub | 채널별 DB 입력 시간 조회 |
| Sheet_조회_INPUT_DB_내역_조회 | Public Sub | Input DB 전체 내역 조회 |

### 3.3 DB 배포 관리 (Sheet_DB_배포.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_DB_배포_Clear | Public Sub | 배포 화면 데이터 초기화 |
| Sheet_DB_배포_Query | Public Sub | 신규 배포 대상 DB 조회 |
| Sheet_DB_배포_추가배포_Query | Public Sub | 추가 배포 대상 DB 조회 |
| Sheet_DB_배포_DB_Upload | Public Sub | TM별 DB 배포 실행 및 회수자 관리 |
| Sheet_DB_배포_회수자_점검 | Public Sub | 과거 담당 TM과 현재 배포 TM 비교 |
| Sheet_DB_배포_구DB_참고자료만_조희 | Public Sub | 구DB 관련 참고 정보 조회 |

### 3.4 Input DB 처리 (Sheet_InputDB_작업.bas, Sheet_InputDB_입력.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_InputDB_작업_Clear | Public Sub | 작업 시트 초기화 |
| Sheet_InputDB_작업_PreProcessing | Public Sub | 다양한 소스의 DB를 표준 형식으로 변환 |
| Sheet_InputDB_작업_배포업무_일괄세팅하기 | Public Sub | 배포 업무 일괄 설정 |
| Sheet_InputDB_입력_Clear | Public Sub | 입력 시트 초기화 |
| Sheet_InputDB_입력_DB_Upload | Public Sub | 표준화된 DB를 시스템에 입력 |

**지원 소스**: 애드인, 방송, 홈페이지_1/2, 모두닥(예약/전화/상담), 모두닥할인몰, 타불라, 페이스북, 지오엔, 빅크레프트, 강남언니

### 3.5 DentWeb 연동 (Sheet_DentWeb_연동.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_DentWeb_연동_Clear | Public Sub | 연동 화면 초기화 |
| Sheet_DentWeb_연동_Query | Public Sub | DentWeb에서 예약 데이터 조회 (신규/취소/누락 내역) |
| Sheet_DentWeb_연동_DB_Upload | Public Sub | 조회한 예약 데이터를 CRM DB에 업로드 |

### 3.6 통계 및 분석 (Sheet_통계.bas, Sheet_DashBoard.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_통계_조회 | Public Sub | 통계 타입별 분기 처리 |
| Sheet_통계_내원환자_리스트_내원일_조회 | Public Sub | 내원일 기준 환자 리스트 |
| Sheet_통계_내원환자_리스트_배포일_조회 | Public Sub | 배포일 기준 환자 리스트 |
| Sheet_통계_TM별_일일_통계_조회 | Public Sub | TM별 일일 성과 통계 |
| Sheet_통계_TM별_DB업체별_Performance_조회 | Public Sub | TM별 DB업체별 성과 분석 |
| Sheet_DashBoard_Query | Public Sub | 대시보드 종합 지표 조회 |
| Sheet_DashBoard_Ref_Date_UP | Public Sub | 대시보드 날짜 증가 |
| Sheet_DashBoard_Ref_Date_Down | Public Sub | 대시보드 날짜 감소 |

### 3.7 구DB 작업 (Sheet_구DB_작업.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_구DB_작업_Clear | Public Sub | 구DB 작업 화면 초기화 |
| Sheet_구DB_작업_Query | Public Sub | 과거 DB 중 재활용 가능한 DB 조회 |
| Sheet_구DB_작업_DB배포시트_옮기기 | Public Sub | 선택한 구DB를 배포 시트로 이동 |
| Sheet_구DB_결번정보_DB_Upload | Public Sub | 결번(응답없음) 정보 업로드 |

### 3.8 예약 및 내원 관리 (Sheet_예약내원정보.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_예약내원정보_Clear | Public Sub | 예약/내원 화면 초기화 |
| Sheet_예약내원정보_Query | Public Sub | 예약 및 내원 정보 조회 |
| Sheet_예약내원정보_DB_Update | Public Sub | 내원 여부 업데이트 |

### 3.9 DB Agent 클래스 (DB_Agent.cls)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Connect_DB | Public Function | 데이터베이스 연결 |
| Close_DB | Private Function | 데이터베이스 연결 종료 |
| Select_DB | Function | SELECT 쿼리 실행 |
| Insert_update_DB | Function | INSERT/UPDATE/DELETE 쿼리 실행 |
| Begin_Trans | Public Function | 트랜잭션 시작 |
| Commit_Trans | Public Function | 트랜잭션 커밋 |
| Rollback_Trans | Public Function | 트랜잭션 롤백 |

### 3.10 유틸리티 함수 (Util.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| make_SQL | Function | SQL 파라미터 바인딩 (:param01, :param02 등) |
| Cnvt_to_Date | Public Function | yyyyMMdd → yyyy-MM-dd 변환 |
| ADD_Date | Public Function | 영업일 기준 날짜 계산 |
| File_Exists | Public Function | 파일/디렉토리 존재 확인 |
| Get_Row_by_Find | Public Function | 시트에서 특정 값 찾기 |
| Get_Row_num | Public Function | 데이터 행 개수 세기 |

### 3.11 DB 업체 결제 관리 (Sheet_DB업체_결제.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_DB업체_결제_Clear | Public Sub | 결제 화면 초기화 |
| Sheet_DB업체_결제_Query | Public Sub | DB 업체별 결제 정보 조회 |
| Run_SQL_C77_동적IN_출력 | Public Sub | 동적 IN 절을 사용한 SQL 실행 |

### 3.12 SQL 실행기 (Sheet_SQL_실행기.bas)

| 함수명 | 유형 | 주요 역할 |
|--------|------|-----------|
| Sheet_SQL_실행기_Clear | Public Sub | SQL 실행 화면 초기화 |
| Sheet_SQL_실행기_Query | Public Sub | 사용자 정의 SQL 실행 |

---

## 4. 주요 상수 및 전역 변수

### 4.1 시트별 상수

| 시트명 | 상수명 | 값 | 설명 |
|--------|--------|-----|------|
| 조회 | This_Sheet_Name | "조회" | 현재 시트 이름 |
| 조회 | START_ROW_NUM | 8 | 데이터 시작 행 |
| 조회 | MAX_ROW_NUM | 10000 | 최대 행 수 |
| DB_배포 | This_Sheet_Name | "DB_배포" | 현재 시트 이름 |
| DB_배포 | START_ROW_NUM | 11 | 데이터 시작 행 |
| DB_배포 | MAX_ROW_NUM | 5000 | 최대 행 수 |
| InputDB_작업 | This_Sheet_Name | "InputDB_작업" | 현재 시트 이름 |
| InputDB_작업 | Target_Sheet_Name | "InputDB_입력" | 타겟 시트 이름 |
| InputDB_작업 | START_ROW_NUM | 19 | 데이터 시작 행 |
| DentWeb_연동 | START_ROW_NUM | 8 | 데이터 시작 행 |
| DentWeb_연동 | MAX_ROW_NUM | 10000 | 최대 행 수 |
| Oracle_조회 | MAX_ROW_NUM | 100000 | 최대 행 수 (대용량) |

### 4.2 DB Agent 클래스 변수

| 변수명 | 타입 | 설명 |
|--------|------|------|
| connection | ADODB.connection | 데이터베이스 연결 객체 |
| connect_str | String | 연결 문자열 |
| sql_str | String | SQL 쿼리 문자열 |
| result_recordset | ADODB.Recordset | 쿼리 결과 레코드셋 |

---

## 5. DB 스키마 정보

### 5.1 주요 테이블

#### INPUT_DB (입력 DB 테이블)
```
ref_date          - 기준일자 (yyyyMMdd)
ref_seq           - 입력 회차
DB_SRC_1          - DB 출처 1 (채널 구분)
DB_SRC_2          - DB 출처 2 (세부 구분)
event_type        - 이벤트 타입
client_name       - 고객명
phone_no          - 전화번호
SEQ_NO            - 동일 DB 내 시퀀스
db_input_date     - DB 입력일
db_input_time     - DB 입력시간
age               - 나이
gender            - 성별
db_memo_1         - 메모 1
db_memo_2         - 메모 2
db_memo_3         - 메모 3
qual_mng          - MNG 품질 (당일중복 등)
qual_tm           - TM 품질
DUPL_CNT_1        - 중복 횟수 1
DUPL_CNT_2        - 중복 횟수 2
DUPL_LAST_DATE_1  - 마지막 중복 일자 1
DUPL_LAST_DATE_2  - 마지막 중복 일자 2
DUPL_LAST_DB_SRC_1 - 마지막 중복 DB 출처
```

#### DB_DIST_HIS (DB 배포 이력)
```
REF_DATE           - 배포 기준일자
REF_SEQ            - 배포 회차
DB_SRC_1           - DB 출처 1
DB_SRC_2           - DB 출처 2
PHONE_NO           - 전화번호
ASSIGNED_TM_NO     - 배정 TM 번호
ASSIGNED_TM_NAME   - 배정 TM 이름
CLIENT_NAME        - 고객명
EVENT_TYPE         - 이벤트 타입
DB_INPUT_DATE      - DB 입력일
DB_INPUT_TIME      - DB 입력시간
DB_MEMO_1          - 메모
ASSIGNED_STATUS    - 배정 상태 (배포완료/배포예정/배포안함)
INPUT_DB_QUAL      - Input DB 품질
IS_OLD_DB          - 구DB 여부 (Y/N)
MNG_MEMO           - MNG 메모
INPUTDB_REF_DATE   - Input DB 기준일자
INPUTDB_REF_SEQ    - Input DB 회차
INPUTDB_SEQ_NO     - Input DB 시퀀스
ALT_USER_NO        - 작업자 번호
```

#### TM_CALL_LOG (TM 콜 로그)
```
ref_date           - 기준일자
phone_no           - 전화번호
TM_NO              - TM 번호
TM_NAME            - TM 이름
call_result        - 통화 결과
call_log           - 통화 내용
reservation_date   - 예약일자
reservation_status - 예약 상태
visit_status       - 내원 여부
visit_date         - 내원일자
```

#### DB_WITHDRAW_HIS (DB 회수 이력)
```
REF_DATE           - 기준일자
PHONE_NO           - 전화번호
TM_NO              - 회수 TM 번호
TM_NAME            - 회수 TM 이름
CLIENT_NAME        - 고객명
DB_SRC_1           - DB 출처 1
DB_SRC_2           - DB 출처 2
ALT_USER_NO        - 작업자 번호
```

#### DENTWEB 관련 테이블
```
NID                - 예약 ID
예약날짜            - 예약 날짜
예약시각            - 예약 시간
작성날짜            - 작성 날짜
작성시각            - 작성 시간
N환자ID            - 환자 ID
환자이름            - 환자 이름
환자전화번호        - 전화번호
차트번호            - 차트 번호
N소요시간          - 소요 시간
N예약종류          - 예약 종류
N이행현황          - 이행 현황 (미도래/이행/취소)
N담당의사          - 담당 의사
N담당직원          - 담당 직원
SZ예약내용        - 예약 내용
SZ메모            - 메모
T최종수정날짜      - 최종 수정 날짜
T최종수정시각      - 최종 수정 시간
```

#### MNG_결번_관리 (결번 관리)
```
phone_no           - 전화번호
client_name        - 고객명
note               - 비고
alt_date           - 등록일자
```

#### CHART_SALES_STATUS (차트 매출 현황)
```
INPUT_DATE         - 입력일
RESERV_DATE        - 예약일
RESERV_TIME        - 예약 시간
CHART_NO           - 차트 번호
CLIENT_NAME        - 고객명
PHONE_NO           - 전화번호
EVENT_TYPE         - 이벤트 타입
DB_SRC_2           - DB 출처
REF_DATE           - 기준일자
DB_INPUT_DATE      - DB 입력일
TM_NAME            - TM 이름
CONSENT_SALES      - 동의 매출
```

#### AD_COST_COMPANY (광고비 현황)
```
DB_INPUT_DATE      - DB 입력일
DB_SRC_1           - DB 출처 1
DB_SRC_2           - DB 출처 2
EVENT              - 이벤트
AD_COST            - 광고비
```

#### USER_INFO (사용자 정보)
```
USER_NO            - 사용자 번호
USER_NAME          - 사용자 이름
USER_PASS          - 비밀번호
```

---

### 5.2 주요 SQL 쿼리 패턴

**파라미터 바인딩 패턴**:
```vba
':param01, :param02, ... :paramNN
```

**날짜 포맷**:
- 저장: `yyyyMMdd` (예: 20250817)
- 표시: `yyyy-MM-dd` (예: 2025-08-17)

---

## 6. 핵심 비즈니스 로직

### 6.1 DB 배포 프로세스

```
1. Input DB 입력
   └─ 다양한 채널(홈페이지, 모두닥, 방송 등)에서 DB 수집
   └─ 표준 형식으로 변환 (Sheet_InputDB_작업_PreProcessing)
   └─ 중복 체크 및 품질 검증
   └─ INPUT_DB 테이블에 저장

2. DB 배포 대상 조회
   └─ 미배포 DB 조회 (Sheet_DB_배포_Query)
   └─ 과거 콜 기록 확인
   └─ DentWeb 예약 상황 확인
   └─ 결번 정보 확인
   └─ 회수자 정보 확인

3. TM 배정
   └─ 배포 대상 TM 선택
   └─ 회수자 점검 (과거 담당자와 다른 경우)
   └─ 중복 배정 체크

4. DB 배포 실행
   └─ DB_DIST_HIS 테이블에 배포 이력 저장
   └─ 회수자 정보 저장 (DB_WITHDRAW_HIS)
   └─ 결번 정보 저장 (MNG_결번_관리)
```

### 6.2 DentWeb 연동 프로세스

```
1. DentWeb 데이터 조회
   └─ CRM의 MAX NID 확인
   └─ DentWeb에서 신규 예약 조회
   └─ 누락 내역 조회 (NID 중간 누락 확인)

2. 데이터 분류
   ├─ 신규내역: CRM에 없는 예약
   ├─ 당일내역: 당일 생성된 예약
   ├─ 취소내역: 취소된 예약
   └─ 누락내역: NID 중간 누락 예약

3. CRM DB 업로드
   └─ 신규/당일내역: INSERT
   └─ 취소내역: 기존 DELETE 후 INSERT
   └─ 누락내역: 중복 체크 후 INSERT
```

### 6.3 통계 생성 로직

```
1. 내원 환자 리스트
   └─ 내원일 또는 배포일 기준 조회
   └─ TM별, DB업체별 그룹핑

2. TM별 일일 통계
   └─ 배포 DB 수
   └─ 콜 수행률
   └─ 예약 건수
   └─ 내원 건수
   └─ 전환율 계산

3. DB업체별 Performance
   └─ 업체별 배포 수
   └─ 업체별 예약률
   └─ 업체별 내원률
   └─ ROI 계산 (광고비 대비 매출)
```

### 6.4 구DB 재활용 프로세스

```
1. 재활용 가능 DB 조건
   ├─ 과거 배포 후 일정 기간 경과
   ├─ 콜 미수행 또는 특정 결과
   ├─ 예약 미완료
   └─ 결번 아님

2. 구DB 작업
   └─ 조건에 맞는 DB 조회
   └─ 재배포 여부 판단
   └─ 선택한 DB를 배포 시트로 이동
```

---

## 7. 다른 파일과의 차별점 (Marketing vs Marketing_an)

### 7.1 파일 크기 및 복잡도

| 항목 | Marketing.vba | Marketing_an.vba |
|------|---------------|------------------|
| 총 라인 수 | 10,717 lines | 470 lines |
| 모듈 수 | 20+ 모듈 | 1개 모듈 (Module1.bas) |
| 클래스 수 | 2개 (DB_Agent, ThisWorkbook) | 0개 |
| 함수/프로시저 수 | 80+ 개 | 6개 |

### 7.2 기능 차이

#### Marketing.vba (주 파일)
**목적**: 종합 CRM 관리 시스템

**주요 기능**:
1. **DB 입력 및 배포 관리**
   - 다양한 소스(홈페이지, 모두닥, 방송 등)의 DB 처리
   - TM별 DB 배포 및 회수 관리
   - 중복 체크 및 품질 관리

2. **DentWeb 연동**
   - 예약 시스템과 실시간 동기화
   - 예약 상태 추적 (미도래/이행/취소)

3. **통계 및 분석**
   - TM별 성과 분석
   - DB업체별 ROI 분석
   - 일일/월간 통계

4. **사용자 관리**
   - 로그인 시스템
   - 권한 관리
   - 작업 이력 추적

5. **고급 기능**
   - 구DB 재활용
   - 결번 관리
   - 회수자 자동 추적
   - 트랜잭션 관리

#### Marketing_an.vba (분석 파일)
**목적**: 광고비 및 매출 데이터 분석

**주요 기능**:
1. **광고비 데이터 입력**
   - `AddCostInsert()`: AD_COST_COMPANY 테이블에 광고비 데이터 삽입
   - 단순 반복문으로 Excel 데이터를 DB에 INSERT

2. **매출 데이터 입력**
   - `SalesDataInsert()`: CHART_SALES_STATUS 테이블에 매출 데이터 삽입
   - 차트 번호별 매출 현황 저장

3. **데이터 조회**
   - `RunQueryAndLoadResult()`: 광고비 현황 조회
   - `RunQuery_Load_Work()`: 작업 데이터 조회
   - `RunQuery_ChartSalesByChartNo()`: 차트 번호별 매출 조회

4. **데이터 삭제**
   - `Run_Delete_ChartSalesStatus()`: 매출 데이터 초기화
   - `Run_Delete_AdCostCompany()`: 광고비 데이터 초기화

### 7.3 코드 구조 차이

#### Marketing.vba
```
복잡한 모듈 구조
├─ DB_Agent 클래스 (연결 풀링, 트랜잭션)
├─ 시트별 모듈 (20+ 개)
├─ 유틸리티 모듈
├─ SQL 래퍼
└─ 상태 관리 및 검증 로직
```

#### Marketing_an.vba
```
단순한 스크립트 구조
├─ Module1만 존재
├─ 직접 ADODB 연결
├─ 단순 INSERT/SELECT/DELETE만 수행
└─ 유효성 검증 최소화
```

### 7.4 데이터베이스 접근 방식

#### Marketing.vba
```vba
' DB_Agent 클래스 사용
Dim select_db_agent As DB_Agent
Set select_db_agent = New DB_Agent
select_db_agent.Begin_Trans
' ... 작업 수행
select_db_agent.Commit_Trans
```
- 클래스 기반 추상화
- 트랜잭션 관리
- 에러 핸들링
- 연결 자동 관리

#### Marketing_an.vba
```vba
' 직접 ADODB 사용
Dim Connection As Object
Set Connection = CreateObject("ADODB.Connection")
Connection.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
Connection.Execute sql
Connection.Execute "COMMIT"
Connection.Close
```
- 직접 연결 관리
- 수동 커밋
- 단순 에러 처리

### 7.5 대상 사용자

| 항목 | Marketing.vba | Marketing_an.vba |
|------|---------------|------------------|
| 대상 사용자 | TM, 관리자, 마케팅 담당자 | 데이터 분석가, 경영진 |
| 사용 빈도 | 일일 사용 (High) | 주간/월간 분석 (Medium) |
| 데이터 볼륨 | 대용량 (10,000+ rows/day) | 중간 규모 (aggregated data) |
| 실시간성 | 실시간 처리 필요 | 배치 처리 가능 |

### 7.6 핵심 차이점 요약

1. **복잡도**
   - Marketing: 엔터프라이즈급 CRM 시스템
   - Marketing_an: 간단한 데이터 입출력 도구

2. **기능 범위**
   - Marketing: DB 전체 라이프사이클 관리
   - Marketing_an: 광고비/매출 데이터 관리만

3. **데이터 품질**
   - Marketing: 중복 체크, 품질 검증, 회수자 관리
   - Marketing_an: 기본적인 데이터 입력만

4. **통합성**
   - Marketing: DentWeb 연동, 타 시스템 연계
   - Marketing_an: 독립적인 분석 도구

5. **사용 목적**
   - Marketing: 일상적인 운영 업무
   - Marketing_an: 성과 분석 및 의사결정 지원

### 7.7 공통점

1. 동일한 데이터베이스 사용 (knock_crm_real)
2. ADODB를 통한 Oracle/MSSQL 연결
3. Excel VBA 기반
4. 파라미터 바인딩 방식 사용 (:param01, :param02)

---

## 8. 보안 주의사항

**경고**: 이 파일에는 하드코딩된 데이터베이스 자격 증명이 포함되어 있습니다:

```vba
' 메인 DB
DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34

' DentWeb DB
DSN=dentweb;uid=sa;pwd=Q3xzJiwpv2zC
```

**권장 사항**:
1. 자격 증명을 환경 변수나 암호화된 설정 파일로 이동
2. 최소 권한 원칙 적용
3. 정기적인 비밀번호 변경
4. 접근 로그 모니터링

---

## 9. 결론

`Knock_CRM_mng_v.1.60_Marketing.vba`는 치과 CRM 시스템의 핵심 마케팅 DB 관리 도구로, 다음과 같은 특징을 가집니다:

1. **포괄적인 기능**: DB 입력부터 배포, 통계까지 전체 라이프사이클 관리
2. **다중 채널 지원**: 10+ 개의 마케팅 채널 데이터 통합 처리
3. **실시간 연동**: DentWeb 예약 시스템과 실시간 동기화
4. **품질 관리**: 중복 체크, 결번 관리, 회수자 추적 등
5. **성과 분석**: TM별, 업체별 상세 통계 및 ROI 분석

반면, `Marketing_an.vba`는 광고비 및 매출 데이터의 간단한 입출력에 특화된 보조 도구입니다.

두 파일은 상호 보완적인 관계로, Marketing.vba가 일상적인 운영을 담당하고, Marketing_an.vba가 경영 분석을 지원하는 구조입니다.
