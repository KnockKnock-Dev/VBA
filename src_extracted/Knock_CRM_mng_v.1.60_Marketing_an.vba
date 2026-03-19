olevba 0.60.2 on Python 3.11.7 - http://decalage.info/python/oletools
===============================================================================
FILE: Knock_CRM_mng_v.1.60_Marketing_an.xlsm
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO 현재_통합_문서.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/현재_통합_문서'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet3.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module1.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

'Private Connection As ADODB.Connection

Sub AddCostInsert()

    ThisWorkbook.Sheets("업체별 광고비 현황").Activate
    
    Dim Connection As Object
    Dim sql As String
    Dim lastRow As Long
    Dim i As Long
    
    ' ADODB.Connection 객체 생성
    Set Connection = CreateObject("ADODB.Connection")
    Connection.ConnectionTimeout = 300 ' 300초 = 5분
    Connection.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 반복문으로 2번째 행부터 마지막 행까지 데이터 삽입
    For i = 2 To lastRow
        Dim formattedDate As String
        
        sql = "INSERT INTO ad_cost_company (DB_INPUT_DATE, DB_SRC_1, DB_SRC_2, EVENT, AD_COST) VALUES ('" & _
          Cells(i, 1).Value & "', '" & _
          Cells(i, 2).Value & "', '" & _
          Cells(i, 3).Value & "', '" & _
          Cells(i, 4).Value & "', '" & _
          Cells(i, 5).Value & "')"
          
        Connection.Execute sql
        
        Debug.Print i&; "  " & sql
    Next i
    
    ' 변경사항 커밋 (명시적 커밋이 필요한 경우)
    Connection.Execute "COMMIT"
    
    ' 연결 종료
    Connection.Close
    Set Connection = Nothing
    
    MsgBox "데이터가 정상적으로 삽입되었습니다."
End Sub

Sub SalesDataInsert()

    ThisWorkbook.Sheets("차트번호별 매출현황").Activate
    
    Dim Connection As Object
    Dim sql As String
    Dim lastRow As Long
    Dim i As Long
    
    ' ADODB.Connection 객체 생성
    Set Connection = CreateObject("ADODB.Connection")
    Connection.ConnectionTimeout = 300 ' 300초 = 5분
    Connection.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 반복문으로 2번째 행부터 마지막 행까지 데이터 삽입
    For i = 2 To lastRow
        Dim formattedDate As String
        formattedDate = Format(Cells(i, 1).Value, "yyyymmdd")
        'formattedDate2 = Format(Cells(i, 2).Value, "yyyymmdd")
        'formattedDate3 = Format(Cells(i, 9).Value, "yyyymmdd")
        'formattedDate4 = Format(Cells(i, 10).Value, "yyyymmdd")
        ' 각 행에서 A~E열 값 읽기, 문자열이면 작은따옴표 처리 후 SQL 문 생성
        sql = "INSERT INTO CHART_SALES_STATUS (INPUT_DATE, RESERV_DATE, RESERV_TIME, CHART_NO, CLIENT_NAME, PHONE_NO, EVENT_TYPE, DB_SRC_2, REF_DATE, DB_INPUT_DATE, TM_NAME, CONSENT_SALES) VALUES ('" & _
          formattedDate & "', '" & _
          Cells(i, 2).Value & "', '" & _
          Cells(i, 3).Value & "', '" & _
          Cells(i, 4).Value & "', '" & _
          Cells(i, 5).Value & "', '" & _
          Cells(i, 6).Value & "', '" & _
          Cells(i, 7).Value & "', '" & _
          Cells(i, 8).Value & "', '" & _
          Cells(i, 9).Value & "', '" & _
          Cells(i, 10).Value & "', '" & _
          Cells(i, 11).Value & "', '" & _
          Cells(i, 17).Value & "')"
          
        Connection.Execute sql
        
        Debug.Print i&; "  " & sql
    Next i
    
    ' 변경사항 커밋 (명시적 커밋이 필요한 경우)
    Connection.Execute "COMMIT"
    
    ' 연결 종료
    Connection.Close
    Set Connection = Nothing
    
    MsgBox "데이터가 정상적으로 삽입되었습니다."
End Sub

Sub RunQueryAndLoadResult()

    Dim conn As Object
    Dim rs As Object
    Dim sqlText As String
    Dim ws As Worksheet
    Dim targetWs As Worksheet
    Dim param01 As String, param02 As String
    Dim i As Long
    
    '--- 시트 지정 ---
    Set ws = ThisWorkbook.Sheets("SQL")                 ' 쿼리 저장 시트
    Set targetWs = ThisWorkbook.Sheets("업체별 광고비 현황")  ' 파라미터+결과 시트
    
    '--- 파라미터 읽기 ---
    param01 = Trim(targetWs.Range("H2").Text)
    param02 = Trim(targetWs.Range("H3").Text)
    
    If Len(param01) = 0 Or Len(param02) = 0 Then
        MsgBox "업체별 광고비 현황 시트의 H2/H3에 값이 없습니다.", vbCritical
        Exit Sub
    End If
    
    '--- SQL 불러오기 ---
    sqlText = ws.Range("C2").Value
    
    If Right$(Trim$(sqlText), 1) = ";" Then
        sqlText = Left$(Trim$(sqlText), Len(Trim$(sqlText)) - 1)
    End If
    
    sqlText = Replace(sqlText, ":param01", "'" & param01 & "'")
    sqlText = Replace(sqlText, ":param02", "'" & param02 & "'")
    
    '--- DB 연결 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    
    '--- Recordset 실행 ---
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlText, conn
    
    '--- 결과 출력 (컬럼명 + 데이터) ---
    targetWs.Range("A1:E1000000").ClearContents
    
    ' 1행에 컬럼명 출력
    For i = 0 To rs.Fields.Count - 1
        targetWs.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    
    ' 2행부터 데이터 출력
    If Not rs.EOF Then
        targetWs.Range("A2").CopyFromRecordset rs
    End If
    
    targetWs.Columns("A:E").AutoFit
    
    '--- 정리 ---
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    MsgBox "쿼리 실행 완료!", vbInformation

End Sub

Sub RunQuery_Load_Work()

    Dim conn As Object
    Dim rs As Object
    Dim sqlText As String
    Dim wsSQL As Worksheet
    Dim wsOut As Worksheet
    Dim param01 As String, param02 As String
    Dim i As Long, colCount As Long, maxCols As Long

    '--- 시트 지정 ---
    Set wsSQL = ThisWorkbook.Sheets("SQL")     ' 쿼리 저장 시트 (C3)
    Set wsOut = ThisWorkbook.Sheets("작업")     ' 파라미터+결과 시트

    '--- 파라미터 읽기 (작업!S2, S3) ---
    param01 = Trim(wsOut.Range("S2").Text)
    param02 = Trim(wsOut.Range("S3").Text)

    If Len(param01) = 0 Or Len(param02) = 0 Then
        MsgBox "작업 시트의 S2/S3에 날짜 값을 입력하세요 (예: 20250817).", vbCritical
        Exit Sub
    End If

    '--- SQL 불러오기 (SQL!C3) ---
    sqlText = wsSQL.Range("C3").Value

    ' 끝의 세미콜론 제거(있다면)
    If Right$(Trim$(sqlText), 1) = ";" Then
        sqlText = Left$(sqlText, Len(sqlText) - 1)
    End If

    ' 바인드 변수 치환
    sqlText = Replace(sqlText, ":param01", "'" & param01 & "'")
    sqlText = Replace(sqlText, ":param02", "'" & param02 & "'")

    '--- DB 연결 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 300
    conn.CommandTimeout = 300
    conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"

    '--- Recordset 실행 ---
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open sqlText, conn

    '--- 결과 출력 (컬럼명 A2, 데이터 A3~) ---
    maxCols = 9 ' A~I 열까지만 (9컬럼)

    ' A2:I 구간만 클리어
    wsOut.Range("A2:I1000000").ClearContents

    ' 필드 수 (최대 9까지만 반영)
    colCount = rs.Fields.Count
    If colCount > maxCols Then colCount = maxCols

    ' 1) 컬럼명 (A2~I2)
    For i = 0 To colCount - 1
        wsOut.Cells(2, i + 1).Value = rs.Fields(i).Name
    Next i

    ' 2) 데이터 (A3~I)
    If Not rs.EOF Then
        wsOut.Range("A3").CopyFromRecordset rs
    End If

    ' A~I까지만 자동너비
    wsOut.Range("A:I").EntireColumn.AutoFit

    '--- 정리 ---
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "쿼리 실행 완료!", vbInformation

End Sub

Sub RunQuery_ChartSalesByChartNo()

    Dim conn As Object
    Dim rs As Object
    Dim sqlText As String
    Dim wsSQL As Worksheet
    Dim wsOut As Worksheet
    Dim param01 As String, param02 As String
    Dim i As Long, colCount As Long, maxCols As Long

    '--- 시트 지정 ---
    Set wsSQL = ThisWorkbook.Sheets("SQL")                    ' 쿼리 저장 시트 (C4)
    Set wsOut = ThisWorkbook.Sheets("차트번호별 매출현황")     ' 파라미터 + 결과 시트

    '--- 파라미터 읽기 ---
    param01 = Trim(wsOut.Range("T2").Text)
    param02 = Trim(wsOut.Range("T3").Text)

    If Len(param01) = 0 Or Len(param02) = 0 Then
        MsgBox "차트번호별 매출현황 시트의 T2/T3에 날짜(YYYYMMDD)를 입력하세요.", vbCritical
        Exit Sub
    End If

    '--- SQL 불러오기 (SQL!C4) ---
    sqlText = wsSQL.Range("C4").Value

    ' 끝의 세미콜론 제거(있다면)
    If Right$(Trim$(sqlText), 1) = ";" Then
        sqlText = Left$(sqlText, Len(sqlText) - 1)
    End If

    ' 바인드 변수 치환
    sqlText = Replace(sqlText, ":param01", "'" & param01 & "'")
    sqlText = Replace(sqlText, ":param02", "'" & param02 & "'")

    '--- DB 연결 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 300
    conn.CommandTimeout = 300
    conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"

    '--- Recordset 실행 ---
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open sqlText, conn

    '--- 결과 출력 (B~K 열만) ---
    maxCols = 10 ' B~K → 10개 컬럼

    ' B1:K 영역 클리어
    wsOut.Range("B1:K1000000").ClearContents

    ' 필드 수 (최대 10까지만 출력)
    colCount = rs.Fields.Count
    If colCount > maxCols Then colCount = maxCols

    ' 1) 컬럼명 (B1~K1)
    For i = 0 To colCount - 1
        wsOut.Cells(1, i + 2).Value = rs.Fields(i).Name   ' +2 → B열부터
    Next i

    ' 2) 데이터 (B2~K)
    If Not rs.EOF Then
        wsOut.Cells(2, 2).CopyFromRecordset rs   ' B2부터 출력
    End If

    ' 보기 좋게 B~K까지만 자동너비
    wsOut.Range("B:K").EntireColumn.AutoFit

    '--- 정리 ---
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "쿼리 실행 완료!", vbInformation

End Sub

Sub Run_Delete_ChartSalesStatus()

    Dim conn As Object
    Dim sqlText As String
    Dim ws As Worksheet
    Dim userResponse As VbMsgBoxResult
    
    ' 확인 메시지
    userResponse = MsgBox("CHART_SALES_STATUS 테이블을 초기화 합니다." & vbCrLf & _
                          "계속 진행하시겠습니까?", vbYesNo + vbQuestion, "삭제 확인")
    If userResponse = vbNo Then Exit Sub  ' 취소 시 바로 종료
    
    '--- SQL 시트 C6 읽기 ---
    Set ws = ThisWorkbook.Sheets("SQL")
    sqlText = Trim(ws.Range("C6").Value)
    
    If Len(sqlText) = 0 Then
        MsgBox "SQL!C6에 DELETE 쿼리가 없습니다.", vbCritical
        Exit Sub
    End If
    
    ' 끝의 세미콜론 제거(있다면)
    If Right$(sqlText, 1) = ";" Then
        sqlText = Left$(sqlText, Len(sqlText) - 1)
    End If
    
    '--- DB 연결 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 300
    conn.CommandTimeout = 300
    conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    
    conn.Execute sqlText
    conn.Execute "COMMIT"
    
    conn.Close
    Set conn = Nothing
    
    MsgBox "CHART_SALES_STATUS 데이터 삭제 완료!", vbInformation

End Sub


Sub Run_Delete_AdCostCompany()

    Dim conn As Object
    Dim sqlText As String
    Dim ws As Worksheet
    Dim userResponse As VbMsgBoxResult
    
    ' 확인 메시지
    userResponse = MsgBox("AD_COST_COMPANY 테이블을 초기화 합니다." & vbCrLf & _
                          "계속 진행하시겠습니까?", vbYesNo + vbQuestion, "삭제 확인")
    If userResponse = vbNo Then Exit Sub  ' 취소 시 종료
    
    '--- SQL 시트 C5 읽기 ---
    Set ws = ThisWorkbook.Sheets("SQL")
    sqlText = Trim(ws.Range("C5").Value)
    
    If Len(sqlText) = 0 Then
        MsgBox "SQL!C5에 DELETE 쿼리가 없습니다.", vbCritical
        Exit Sub
    End If
    
    ' 끝의 세미콜론 제거(있다면)
    If Right$(sqlText, 1) = ";" Then
        sqlText = Left$(sqlText, Len(sqlText) - 1)
    End If
    
    '--- DB 연결 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 300
    conn.CommandTimeout = 300
    conn.Open "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
    
    conn.Execute sqlText
    conn.Execute "COMMIT"
    
    conn.Close
    Set conn = Nothing
    
    MsgBox "AD_COST_COMPANY 데이터 삭제 완료!", vbInformation

End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet4.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet4'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet5.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet9.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet6.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet6'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet7.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet7'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet8.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet8'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit


