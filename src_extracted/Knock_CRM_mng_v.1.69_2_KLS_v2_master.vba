olevba 0.60.2 on Python 3.11.7 - http://decalage.info/python/oletools
===============================================================================
FILE: Knock_CRM_mng_v.1.69_2_KLS_v2_master.xlsm
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub Workbook_open()

    Call Workbook_Initialize

End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet_조회.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_조회'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "조회"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000



Public Sub Sheet_조회_Query()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    
    Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Selection.Font.Bold = False
        
    Sheets(This_Sheet_Name).cells(6, 3) = ""
    Sheets(This_Sheet_Name).cells(6, 5) = ""
        
    
    Select Case Sheets(This_Sheet_Name).cells(2, 8)
                
        Case "INPUT_DB 채널별 입력시간 조회"
        
            Call Sheet_조회_INPUT_DB_source별_입력시간_조회
        
        Case "DB 배포 내역 조회"
    
            Call Sheet_조회_DB_배포_내역_조회

        Case "INPUT_DB 내역 조회"
        
            Call Sheet_조회_INPUT_DB_내역_조회
    
        Case "Call Log 조회"
        
            Call Sheet_조회_Call_Log_조회
        
        Case "연락처 History 전체 조회"
        
            If Sheets(This_Sheet_Name).cells(3, 3) = "" Then
                
                MsgBox "검색할 전화번호를 입력해주세요. (Cell C3)"
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                
                Range("B8").Select
                
                Exit Sub
            
            End If

            Call Sheet_조회_연락처_history_전체_조회

        Case "DB 회수 내역 조회"
        
            Call Sheet_조회_DB_회수_내역_조회


        Case Else

            MsgBox "ERROR"
        
    End Select
    
    

    Rows((START_ROW_NUM - 1) & ":" & (START_ROW_NUM - 1)).Select
    
    If Sheets(This_Sheet_Name).AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    Selection.AutoFilter
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    MsgBox "Query Done"

End Sub


Public Sub Sheet_조회_DB_회수_내역_조회()

   With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("조회_DB_회수_내역_조회").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        .cells(6, 5) = .cells(2, 3) & " ~ " & .cells(2, 5) & " 의 DB 회수 전체 내역"
        

    End With


End Sub



Public Sub Sheet_조회_Call_Log_조회()

   With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("조회_DB_CALL_LOG_조회").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        .cells(6, 5) = .cells(2, 3) & " ~ " & .cells(2, 5) & " 의 Call Log 전체 내역"
        

    End With


End Sub


Public Sub Sheet_조회_연락처_history_전체_조회()

    With Sheets(This_Sheet_Name)
  
        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        'Dim ref_date As String
        Dim sql_str As String
        'Dim user_no As String

        Dim phone_no As String
        
        
        phone_no = .cells(3, 3)
        row_idx = START_ROW_NUM
        
        'ref_date = Format(.Cells(2, 3), "yyyyMMdd")
        'user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        '----------------------------------------------------------------------------
        .cells(row_idx, 2) = "Input DB"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_1").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        row_idx = row_idx + 2
        
        '-------------------------------------------------------------------------------------------------------------------------------
        
        
        .cells(row_idx, 2) = "DB 배포"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_2").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        row_idx = row_idx + 2
        
        '-------------------------------------------------------------------------------------------------------------------------------
   
       
        .cells(row_idx, 2) = "Call Log"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_3").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "DB 회수"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_4").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "덴트웹(CRM 기준) 정보"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_5").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "MNG_결번 정보"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("조회_연락처_history_조회_6").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, phone_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        
        '-------------------------------------------------------------------------------------------------------------------------------

 
   '     .Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        .cells(6, 5) = .cells(3, 3) & " 에 대한 CRM Data 전체"

    End With



End Sub


Public Sub Sheet_조회_DB_배포_내역_조회()

   With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("조회_DB_배포_내역_조회").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        .cells(6, 5) = .cells(2, 3) & " ~ " & .cells(2, 5) & " 의 DB 배포 내역"
        

    End With


End Sub


Public Sub Sheet_조회_INPUT_DB_source별_입력시간_조회()

    With Sheets(This_Sheet_Name)


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim REF_DATE As String
        Dim sql_str As String
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        REF_DATE = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("조회_INPUT_DB_source별_입력시간_조회").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        
        sql_str = Range("조회_INPUT_DB_source별_입력시간_조회_추가").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        

        .cells(6, 3) = .cells(6, 3) + CInt(select_db_agent.DB_result_recordset.RecordCount)
        .cells(6, 5) = "현재 최근 채널별 INPUT DB 내역"
        
'        While Not select_db_agent.DB_result_recordset.EOF
        
'    ref_date,
'    DB_SRC_1,
'    event_type,
'    DB_SRC_2,
'    client_name,
'    phone_no,
'    SEQ_NO,
'    db_input_date,
'    db_input_time,
'    age,
'    gender,
'    db_memo_1,
'    qual_mng,
'    qual_tm,
'    DUPL_CNT_1,
'    DUPL_CNT_2,
'    DUPL_LAST_DATE_1,
'    DUPL_LAST_DATE_2,
'    DUPL_LAST_DB_SRC_1


'           .Cells(row_idx, 2) = select_db_agent.DB_result_recordset(0)
'           .Cells(row_idx, 3) = select_db_agent.DB_result_recordset(1)
'            .Cells(row_idx, 4) = select_db_agent.DB_result_recordset(2)
'            .Cells(row_idx, 5) = select_db_agent.DB_result_recordset(3)
'            .Cells(row_idx, 6) = select_db_agent.DB_result_recordset(4)
'            .Cells(row_idx, 7) = select_db_agent.DB_result_recordset(5)
'            .Cells(row_idx, 8) = select_db_agent.DB_result_recordset(6)
''            .Cells(row_idx, 9) = select_db_agent.DB_result_recordset(7)
          
'            select_db_agent.DB_result_recordset.MoveNext
        
'            row_idx = row_idx + 1
        
'        Wend

    End With


'    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
'        Rows("8:8").Select
'        Selection.AutoFilter
'    End If



'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
    


End Sub



Public Sub Sheet_조회_INPUT_DB_내역_조회()

    With Sheets(This_Sheet_Name)


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("조회_INPUT_DB_내역_조회").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        .cells(6, 5) = .cells(2, 3) & " ~ " & .cells(2, 5) & " 의 INPUT DB 내역"
        
'        While Not select_db_agent.DB_result_recordset.EOF
        
'    ref_date,
'    DB_SRC_1,
'    event_type,
'    DB_SRC_2,
'    client_name,
'    phone_no,
'    SEQ_NO,
'    db_input_date,
'    db_input_time,
'    age,
'    gender,
'    db_memo_1,
'    qual_mng,
'    qual_tm,
'    DUPL_CNT_1,
'    DUPL_CNT_2,
'    DUPL_LAST_DATE_1,
'    DUPL_LAST_DATE_2,
'    DUPL_LAST_DB_SRC_1


'           .Cells(row_idx, 2) = select_db_agent.DB_result_recordset(0)
'           .Cells(row_idx, 3) = select_db_agent.DB_result_recordset(1)
'            .Cells(row_idx, 4) = select_db_agent.DB_result_recordset(2)
'            .Cells(row_idx, 5) = select_db_agent.DB_result_recordset(3)
'            .Cells(row_idx, 6) = select_db_agent.DB_result_recordset(4)
'            .Cells(row_idx, 7) = select_db_agent.DB_result_recordset(5)
'            .Cells(row_idx, 8) = select_db_agent.DB_result_recordset(6)
''            .Cells(row_idx, 9) = select_db_agent.DB_result_recordset(7)
          
'            select_db_agent.DB_result_recordset.MoveNext
        
'            row_idx = row_idx + 1
        
'        Wend

    End With

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet_Config.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_Config'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "Config"


Public Sub Sheet_Config_Intialize()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim select_db_agent As DB_Agent
    Dim sql_str As String
    
    sql_str = ""
    
    With Sheets(This_Sheet_Name)
    
    Sheets(This_Sheet_Name).Select
    
        Set select_db_agent = New DB_Agent
                            
        sql_str = Range("Config_Initialize_Code_SELECT_SQL").Offset(0, 0).Value2
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
            
        Range("J18").CopyFromRecordset select_db_agent.DB_result_recordset
        
        
        sql_str = Range("Config_Initialize_USER_SELECT_SQL").Offset(0, 0).Value2
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
            
        Range("B18").CopyFromRecordset select_db_agent.DB_result_recordset
        
        
           
        Set select_db_agent = Nothing
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Done"

End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet_DentWeb_연동.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DentWeb_연동'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DentWeb_연동"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_DentWeb_연동_Clear()

    Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("B8").Select

    MsgBox "Done"

End Sub


Public Sub Sheet_DentWeb_연동_Query()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    
    Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        
    With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        Dim cur_max_nid As Long
        
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""


        '--------------------------------------------------------------------
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("DentWeb_연동_MAX_NID_SELECT_SQL").Offset(0, 0).Value2
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        cur_max_nid = select_db_agent.DB_result_recordset(0)

        '--------------------------------------------------------------------


        sql_str = Range("DentWeb_연동_덴트웹_누락_SELECT_SQL").Offset(0, 0).Value2
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Dim nid_str As String
        nid_str = ""
        Dim nid_temp As Long
        
        While Not select_db_agent.DB_result_recordset.EOF
        
            If select_db_agent.DB_result_recordset(4) = 2 Then
            
                nid_str = nid_str & select_db_agent.DB_result_recordset(5) & ","
            
            ElseIf select_db_agent.DB_result_recordset(4) > 2 Then
            
                For nid_temp = select_db_agent.DB_result_recordset(5) To select_db_agent.DB_result_recordset(6)
                    nid_str = nid_str & nid_temp & ","
                Next
            
            End If
        
            select_db_agent.DB_result_recordset.MoveNext
        
        Wend

        nid_str = Left(nid_str, Len(nid_str) - 1)

        '--------------------------------------------------------------------
     
        Set select_db_agent = New DB_Agent
        select_db_agent.DB_connect_str = "Provider=SQLOLEDB;" & _
                                          "Data Source=192.168.0.245,1436;" & _
                                          "Initial Catalog=DentWeb;" & _
                                          "User ID=kkpt;" & _
                                          "Password=kkpt12#$;"
        
        sql_str = Range("DentWeb_연동_덴트웹_SELECT_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, cur_max_nid, ref_date_1 & "0000", ref_date_2 & "2400", ref_date_1, ref_date_2)
        
        sql_str = Replace(sql_str, ":param06", nid_str)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        '.Cells(6, 5) = .Cells(2, 3) & " ~ " & .Cells(2, 5) & " 의 Call Log 전체 내역"
        
        
        ' 작은 따옴표 삭제
        row_idx = START_ROW_NUM
        While .cells(row_idx, 2) <> ""
            
            .cells(row_idx, 15) = WorksheetFunction.Substitute(.cells(row_idx, 15), "'", " ")
            .cells(row_idx, 16) = WorksheetFunction.Substitute(.cells(row_idx, 16), "'", " ")
         
            row_idx = row_idx + 1
        Wend
        
        

    End With
    
    

    Rows((START_ROW_NUM - 1) & ":" & (START_ROW_NUM - 1)).Select
    
    If Sheets(This_Sheet_Name).AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    Selection.AutoFilter
              
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    MsgBox "Query Done"

End Sub




Public Sub Sheet_DentWeb_연동_DB_Upload()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.cells(2, 3) & "~" & .cells(2, 5) & " 일자 기준으로 아래 data를 DB Upload 하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Dim row_idx As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim user_no As String
        Dim sql_str As String
        Dim 이행현황 As String
        'Dim qual_mng_result As String
                
      
        row_idx = START_ROW_NUM
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""
        
        
                
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        
        'sql_str = Range("DentWeb_연동_덴트웹_DELETE_SQL").Offset(0, 0).Value2
        'sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        'If Not select_db_agent.Insert_update_DB(sql_str) Then
'        '    MsgBox "ERROR"
        'End If
                
                
        '--------------------------------------------------------------------
        
        Dim cur_max_nid As Long
        
        sql_str = Range("DentWeb_연동_MAX_NID_SELECT_SQL").Offset(0, 0).Value2
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        cur_max_nid = select_db_agent.DB_result_recordset(0)

        '--------------------------------------------------------------------
                
        
        While .cells(row_idx, 2) <> ""


            If .cells(row_idx, 2) = "신규내역" Or (.cells(row_idx, 2) = "당일내역" And .cells(row_idx, 3) <= cur_max_nid) Then
            
                        
                If .cells(row_idx, 2) = "당일내역" Then
            
                    sql_str = Range("DentWeb_연동_덴트웹_취소_DELETE_SQL").Offset(0, 0).Value2
                    sql_str = make_SQL(sql_str, .cells(row_idx, 3))
        
                    If Not select_db_agent.Insert_update_DB(sql_str) Then
'                        MsgBox "ERROR"
                    End If
                End If
        
                sql_str = Range("DentWeb_연동_덴트웹_INSERT_SQL").Offset(0, 0).Value2
                
    '    :param01    --    NID,
    ',    :param02    --    예약날짜,
    ',    :param03    --    예약시각,
    ',    :param04    --    작성날짜,
    ',    :param05    --    작성시각,
    ',    :param06    --    N환자ID,
    ',    :param07    --    환자이름,
    ',    :param08    --    환자전화번호,
    ',    :param09    --    차트번호,
    ',    :param10    --    N소요시간,
    ',    :param11    --    N예약종류,
    ',    :param12    --    N이행현황,
    ',    :param13    --    N담당의사,
    ',    :param14    --    N담당직원,
    ',    :param15    --    SZ예약내용,
    ',    :param16    --    SZ메모,
    ',    :param17    --    T최종수정날짜,
    ',    :param18    --    T최종수정시각,
    ',    :param19    --    ALT_USER_NO

                sql_str = make_SQL(sql_str, _
                                    .cells(row_idx, 3), _
                                    Left(.cells(row_idx, 4), 8), _
                                    Mid(.cells(row_idx, 4), 9, 2) & ":" & Mid(.cells(row_idx, 4), 11, 2) & ":00", _
                                    Left(.cells(row_idx, 5), 8), _
                                    Mid(.cells(row_idx, 5), 9, 2) & ":" & Mid(.cells(row_idx, 5), 11, 2) & ":" & Right(.cells(row_idx, 5), 2), _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 7), _
                                    IIf(.cells(row_idx, 8) = "", "_", .cells(row_idx, 8)), _
                                    .cells(row_idx, 9), _
                                    .cells(row_idx, 10), _
                                    .cells(row_idx, 11), _
                                    .cells(row_idx, 12), _
                                    .cells(row_idx, 13), _
                                    .cells(row_idx, 14), _
                                    .cells(row_idx, 15), _
                                    .cells(row_idx, 16), _
                                    Format(.cells(row_idx, 17), "yyyyMMdd"), _
                                    Format(.cells(row_idx, 17), "hh:mm:ss"), _
                                    user_no)
            
                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
            ElseIf .cells(row_idx, 2) = "취소내역" Then
            
                sql_str = Range("DentWeb_연동_덴트웹_취소_DELETE_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 3))
        
                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                    
                sql_str = Range("DentWeb_연동_덴트웹_INSERT_SQL").Offset(0, 0).Value2
                
    '    :param01    --    NID,
    ',    :param02    --    예약날짜,
    ',    :param03    --    예약시각,
    ',    :param04    --    작성날짜,
    ',    :param05    --    작성시각,
    ',    :param06    --    N환자ID,
    ',    :param07    --    환자이름,
    ',    :param08    --    환자전화번호,
    ',    :param09    --    차트번호,
    ',    :param10    --    N소요시간,
    ',    :param11    --    N예약종류,
    ',    :param12    --    N이행현황,
    ',    :param13    --    N담당의사,
    ',    :param14    --    N담당직원,
    ',    :param15    --    SZ예약내용,
    ',    :param16    --    SZ메모,
    ',    :param17    --    T최종수정날짜,
    ',    :param18    --    T최종수정시각,
    ',    :param19    --    ALT_USER_NO

                sql_str = make_SQL(sql_str, _
                                    .cells(row_idx, 3), _
                                    Left(.cells(row_idx, 4), 8), _
                                    Mid(.cells(row_idx, 4), 9, 2) & ":" & Mid(.cells(row_idx, 4), 11, 2) & ":00", _
                                    Left(.cells(row_idx, 5), 8), _
                                    Mid(.cells(row_idx, 5), 9, 2) & ":" & Mid(.cells(row_idx, 5), 11, 2) & ":" & Right(.cells(row_idx, 5), 2), _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 7), _
                                    IIf(.cells(row_idx, 8) = "", "_", .cells(row_idx, 8)), _
                                    .cells(row_idx, 9), _
                                    .cells(row_idx, 10), _
                                    .cells(row_idx, 11), _
                                    .cells(row_idx, 12), _
                                    .cells(row_idx, 13), _
                                    .cells(row_idx, 14), _
                                    .cells(row_idx, 15), _
                                    .cells(row_idx, 16), _
                                    Format(.cells(row_idx, 17), "yyyyMMdd"), _
                                    Format(.cells(row_idx, 17), "hh:mm:ss"), _
                                    user_no)
            
                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            
                'sql_str = Range("DentWeb_연동_덴트웹_취소_SELECT_SQL").Offset(0, 0).Value2
                'sql_str = make_SQL(sql_str, .Cells(row_idx, 3))
        
                'If Not select_db_agent.Select_DB(sql_str) Then
                '    MsgBox "ERROR"
                'End If
                
                '이행현황 = select_db_agent.DB_result_recordset(0)
                
                'If 이행현황 <> 2 Then
                
                 '   sql_str = Range("DentWeb_연동_덴트웹_취소_UPDATE_SQL").Offset(0, 0).Value2
                 '   sql_str = make_SQL(sql_str, .Cells(row_idx, 3), .Cells(row_idx, 12), ref_date_2)
        
                 '   If Not select_db_agent.Select_DB(sql_str) Then
                 '       MsgBox "ERROR"
                 '   End If
                
                 '   .Cells(row_idx, 1) = "취소 당일"
                
                'End If
                
            ElseIf .cells(row_idx, 2) = "누락내역" Then
            
            
                If WorksheetFunction.CountIf(Range("C:C"), .cells(row_idx, 3)) = 1 Then
        
                    sql_str = Range("DentWeb_연동_덴트웹_INSERT_SQL").Offset(0, 0).Value2

                    sql_str = make_SQL(sql_str, _
                                    .cells(row_idx, 3), _
                                    Left(.cells(row_idx, 4), 8), _
                                    Mid(.cells(row_idx, 4), 9, 2) & ":" & Mid(.cells(row_idx, 4), 11, 2) & ":00", _
                                    Left(.cells(row_idx, 5), 8), _
                                    Mid(.cells(row_idx, 5), 9, 2) & ":" & Mid(.cells(row_idx, 5), 11, 2) & ":" & Right(.cells(row_idx, 5), 2), _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 7), _
                                    IIf(.cells(row_idx, 8) = "", "_", .cells(row_idx, 8)), _
                                    .cells(row_idx, 9), _
                                    .cells(row_idx, 10), _
                                    .cells(row_idx, 11), _
                                    .cells(row_idx, 12), _
                                    .cells(row_idx, 13), _
                                    .cells(row_idx, 14), _
                                    .cells(row_idx, 15), _
                                    .cells(row_idx, 16), _
                                    Format(.cells(row_idx, 17), "yyyyMMdd"), _
                                    Format(.cells(row_idx, 17), "hh:mm:ss"), _
                                    user_no)
            
                    If Not select_db_agent.Insert_update_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                
                Else
                    
                    .cells(row_idx, 1) = "중복 insert pass"
                    
                End If
                
            ElseIf .cells(row_idx, 2) = "당일내역" And .cells(row_idx, 3) > cur_max_nid Then
                
                .cells(row_idx, 1) = "신규+당일 pass"
            
            Else
                MsgBox "Error"
                
            End If
                

            row_idx = row_idx + 1
        
        Wend
                
                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "DB upload Done"
        

End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet_DB_결산.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DB_결산'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DB_결산"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_DB_결산_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        'If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = False
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_DB_결산_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("DB_결산 data를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 기존 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Call Sheet_DB_결산_Clear


        Sheets(This_Sheet_Name).Select

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
                
        'user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
            
        sql_str = Range("DB_결산_조회_SQL").Offset(0, 0).Value2
        
        
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        col_i = 0
        
        'While col_i < select_db_agent.DB_result_recordset.Fields.Count
        '    .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
        '    col_i = col_i + 1
        'Wend
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset


    End With
    


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    MsgBox "Query Done"


End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet5.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet11.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet11'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet16.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet16'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO DB_Agent.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/DB_Agent'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private connection As ADODB.connection
Private connect_str As String
Private sql_str As String
Private result_recordset As ADODB.Recordset

Property Get DB_connect_str() As String
    DB_connect_str = str_DBconnect
End Property

Property Let DB_connect_str(param As String)
    connect_str = param
End Property

Property Get DB_SQL_str() As String
    DB_SQL_str = sql_str
End Property

Property Let DB_SQL_str(param As String)
    sql_str = param
End Property

Property Get DB_result_recordset() As ADODB.Recordset
    Set DB_result_recordset = result_recordset
End Property

Public Function Connect_DB() As Boolean

    Set connection = Nothing
    Set connection = New ADODB.connection
    
    If connect_str = "" Then
        MsgBox "connect_str is NOT specified. ERROR"
        Connect_DB = False
    End If
    
    connection.Open connect_str
    
    Connect_DB = True
    
End Function

Private Function Close_DB() As Boolean

    If connection Is Nothing Then
        MsgBox "connection is NULL. ERROR"
        Close_DB = False
    End If
    
    connection.Close

    Close_DB = True
    
End Function

Function Select_DB(Optional param As String) As Boolean

    If param <> "" Then
        sql_str = param
    End If

    If sql_str = "" Then
        MsgBox "SQL_str is NULL. ERROR"
        Select_DB = False
    End If

    If connection Is Nothing Then
        Connect_DB
    End If
    
    Set result_recordset = New ADODB.Recordset
    
    With result_recordset
      .CursorType = adOpenStatic
      .CursorLocation = adUseClient
      .LockType = adLockReadOnly
      .Source = sql_str
      .ActiveConnection = connection
      .Open
    End With
    
    Select_DB = True

End Function


Function Insert_update_DB(Optional param As String) As Boolean

    Dim cnt_result As Integer
    cnt_result = 0

    If param <> "" Then
        sql_str = param
    End If

    If sql_str = "" Then
        MsgBox "SQL_str is NULL. ERROR"
        Insert_update_DB = False
    End If

    If connection Is Nothing Then
        Connect_DB
    End If
        
    connection.Execute sql_str, cnt_result
    
    If Err <> 0 Or cnt_result = 0 Then
        'MsgBox Err.Description & " " & cnt_result & "개의 row가 변경"
        Insert_update_DB = False
        Exit Function
    End If
    
    Insert_update_DB = True

End Function


Private Sub Class_Initialize()
    
    'If Sheets("Config").Cells(2, 20) = "Real_DB" Then
        'connect_str = "DSN=PEBK;uid=prm;pwd=Rnfjddl13#"
    'Else
        'connect_str = "DSN=TPLT;uid=pilot2;pwd=pilot2"
        'connect_str = "DSN=knock_crm_real;uid=knock_crm;pwd=thsehdrnr123$"
    'End If
    
    connect_str = "DSN=knock_crm_real;uid=knock_crm;pwd=kkptcmr!@34"
        
End Sub

Private Sub Class_Terminate()
    
    If Not result_recordset Is Nothing Then
        result_recordset.Close
    End If
    
    If Not connection Is Nothing Then
        Close_DB
    End If
    
End Sub

Public Function Begin_Trans() As Boolean
    connection.BeginTrans
End Function

Public Function Commit_Trans() As Boolean
    connection.CommitTrans
End Function

Public Function Rollback_Trans() As Boolean
    connection.RollbackTrans
End Function


-------------------------------------------------------------------------------
VBA MACRO SQL_Wrapper.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/SQL_Wrapper'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Function make_SQL(sql_str As String, ParamArray arglist() As Variant) As String

    Dim i As Integer
    
    For i = 0 To UBound(arglist())
    
        'Debug.Print i & "  :  " & arglist(i)
        'Debug.Print i, TypeName(arglist(i)), VarType(arglist(i))
    
        If arglist(i) = "ALL" Then
            sql_str = WorksheetFunction.Substitute(sql_str, ":param" & Format(i + 1, "00"), "'%'")
        Else
            sql_str = WorksheetFunction.Substitute(sql_str, ":param" & Format(i + 1, "00"), "'" & arglist(i) & "'")
        End If
    
    Next i
    
    make_SQL = sql_str

End Function


-------------------------------------------------------------------------------
VBA MACRO Sheet_InputDB_작업_old.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_InputDB_작업_old'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "InputDB_작업"
Const Target_Sheet_Name As String = "InputDB_입력"

Const START_ROW_NUM As Integer = 7
Const MAX_ROW_NUM As Integer = 2000


Public Sub Sheet_InputDB_작업_Clear_()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        If MsgBox("아래 sheet data를 clear 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.Clear
       
       
        Range(Range("INPUTDB_TYPE").Offset(0, 0), Range("INPUTDB_TYPE").Offset(10, 0)).Copy
        Range("B7").Select
        ActiveSheet.Paste
        Range("B7").Select
              
               
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Done"

End Sub


Public Sub Sheet_InputDB_작업_PreProcessing_()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        If MsgBox("아래 sheet data를 다음 sheet로 cleansing 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Call Sheet_InputDB_입력_Clear
        
        Sheets(This_Sheet_Name).Select
    
    
        Dim total_data_cnt As Integer
        Dim row_idx As Integer
        Dim target_row_idx As Integer
                
        Dim inputdb_type As String
        
        
        total_data_cnt = .cells(5, 2)
        row_idx = START_ROW_NUM
        inputdb_type = ""
        
        target_row_idx = Sheet_InputDB_입력.START_ROW_NUM
        
        
        'While .Cells(row_idx, 3) <> "" Or total_data_cnt > 0
        While .cells(row_idx, 3) <> ""
        
            'If .Cells(row_idx, 3) = "" Then
            '    Rows(row_idx & ":" & row_idx).Select
            '    Selection.Delete Shift:=xlUp
            'Else
              
                total_data_cnt = total_data_cnt - 1
                
                If .cells(row_idx, 2) <> "" Then
                    inputdb_type = .cells(row_idx, 2)
                Else
                    .cells(row_idx, 2) = inputdb_type
                End If
                
                
                Select Case .cells(row_idx, 2)
                
                
                    Case "ADD_01"
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 2)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 2)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        '
                        
                        Dim temp_idx_1 As Integer
                        Dim temp_idx_2 As Integer
                        Dim temp_idx_3 As Integer
                        Dim temp_idx_4 As Integer

                        temp_idx_1 = InStr(.cells(row_idx, 3), ".")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 3), ".")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 3), " ")
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 3), ":")
                        
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = IIf(IsDate(.Cells(row_idx, 3)), _
                        '                                                            Format(.Cells(row_idx, 3), "yyyyMMdd"), _
                        '                                                            Left(.Cells(row_idx, 3), 4) & Format(Mid(.Cells(row_idx, 3), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 1), "00") & Format(Mid(.Cells(row_idx, 3), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00"))
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = IIf(IsDate(.Cells(row_idx, 3)), _
                        '                                                            Format(.Cells(row_idx, 3), "hh:mm:ss"), _
                        '                                                            Format(Mid(.Cells(row_idx, 3), temp_idx_3 + 1, temp_idx_4 - temp_idx_3 - 1), "00") & ":" & Right(.Cells(row_idx, 3), 2) & ":00")
                         
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Left(.Cells(row_idx, 3), 4) & Format(Mid(.Cells(row_idx, 3), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 1), "00") & Format(Mid(.Cells(row_idx, 3), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(Mid(.Cells(row_idx, 3), temp_idx_3 + 1, temp_idx_4 - temp_idx_3 - 1), "00") & ":" & Right(.Cells(row_idx, 3), 2) & ":00"
                         
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                         
                         
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 9) = .cells(row_idx, 7)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) =
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                        
                        
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 4).Address).Select
                    
                        Selection.Interior.Color = 65535
                
                
                    Case "방송DB"
                        
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 4)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = Left(.cells(row_idx, 5), Len(.cells(row_idx, 5)) - 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = IIf(Mid(.cells(row_idx, 6), 4, 1) = "-", .cells(row_idx, 6), _
                                                                                  "0" & Left(.cells(row_idx, 6), 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4))
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = "_"
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 10) = Left(Right(.cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = WorksheetFunction.Substitute(.cells(row_idx, 7), Chr(10), "  |  ")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =


                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "홈페이지"
                    
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = IIf(.cells(row_idx, 4) = "임플란트" And .cells(row_idx, 3) = "메인 홈페이지", _
                                                                                    "4_1", _
                                                                                    .cells(row_idx, 4))
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 3)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Left(.cells(row_idx, 6), 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 9), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 9), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "모두닥"
                    
                    
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 2)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 3)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 7), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = "_"
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 8) & "  |  " & .cells(row_idx, 9) & "  |  " & .cells(row_idx, 13)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 13).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    
                    Case "페이스북"
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 4)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 3)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Left(.cells(row_idx, 6), 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 8), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 8) & "  |  " & .Cells(row_idx, 9) & "  |  " & .Cells(row_idx, 13)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                        
                        Selection.Interior.Color = 65535
                    
                    
                    
                    Case "지오엔"
                    
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 4)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 3)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 16), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 16), "hh:mm:ss")
                        ' 나이
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 9) = .cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8) & "  |  " & .cells(row_idx, 11) & "  |  " & .cells(row_idx, 13)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 16).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 11).Address & "," & _
                                .cells(row_idx, 13).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "빅크레프트"
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 4)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 3)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Left(.cells(row_idx, 6), 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 13), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 13), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 8) & "  |  " & WorksheetFunction.Substitute(.cells(row_idx, 10), Chr(10), "  |  ")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 13).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 10).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "강남언니"
                    
                    
                    
                    Case Else

                        MsgBox "ERROR : Unknown Input Source"
        
                    End Select
                
                row_idx = row_idx + 1
                target_row_idx = target_row_idx + 1
                
            'End If
        
        Wend
        
        
        '-----------------------------------------------------------------------------
        
        Sheets(Target_Sheet_Name).Activate
        
        target_row_idx = Sheet_InputDB_입력.START_ROW_NUM
        
        While Sheets(Target_Sheet_Name).cells(target_row_idx, 2) <> ""
        
        
            'If WorksheetFunction.CountIf(Range("E:E"), "=" & Range("E" & target_row_idx)) > 1 Then
            
            '    Rows(target_row_idx & ":" & target_row_idx).Select
                
            '    With Selection.Interior
            '        .Pattern = xlSolid
            '        .PatternColorIndex = xlAutomatic
            '        .ThemeColor = xlThemeColorAccent2
            '        .TintAndShade = 0.799981688894314
            '        .PatternTintAndShade = 0
            '    End With
            
            'End If
            
        
        
            cells(target_row_idx, 1).Formula = "=COUNTIF(F:F,""="" & F" & target_row_idx & ")+" & "IF(E" & target_row_idx & "<>""-"",COUNTIF(E:E,""="" & E" & target_row_idx & "),1)" & "-2"
            
            cells(target_row_idx, 1).Calculate
            
            
            If cells(target_row_idx, 1) > 0 Then
            
                Rows(target_row_idx & ":" & target_row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            End If
            
            target_row_idx = target_row_idx + 1
        
        Wend
        
        
        
        
        
'        Range("C" & START_ROW_NUM & ":" & "J" & (.Cells(5, 2) + START_ROW_NUM - 1)).Select
        'Selection.ClearFormats
        
        
       ' ClearHyperlinks 하이퍼링크 지우기
       ' ClearOutline 윤곽선 지우기
       ' ClearFormat 서식 지우기
        
'        Columns("B:G").Select
'        With Selection.Font
'            .Size = 10
'        End With
'        With Selection
'            .HorizontalAlignment = xlCenter
'        End With
    
    
    
        If Not Sheets(Target_Sheet_Name).AutoFilterMode Then
            Rows("6:6").Select
            Selection.AutoFilter
        End If

        
        Range("B7").Select
               
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Done"

End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet6.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet6'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Workbook_open.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Workbook_open'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Sub Workbook_Initialize()

    ' workbook 전체에 대한 intialize하기
    

    Dim emp_no As String
    Dim emp_pass As String
    Dim emp_name As String
    
    Dim sql_str As String
        
    Dim try_no As Integer


    Dim select_db_agent As DB_Agent

        
    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False
        
    try_no = 0
        
    Set select_db_agent = New DB_Agent
        
    While try_no < 5 And try_no <> -1
        
        emp_no = CStr(InputBox("User ID를 입력해주세요.", "User ID"))
        emp_pass = CStr(InputBox("User Pass를 입력해주세요.", "User Pass"))
    
        sql_str = Range("Login_Select_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, emp_no, emp_pass)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        If select_db_agent.DB_result_recordset.EOF Then
            try_no = try_no + 1
            MsgBox "[Login 실패] 다시 입력해주세요. 총 실패회수 : " & try_no & "번"
        Else
            emp_name = select_db_agent.DB_result_recordset(0)
            try_no = -1
        End If
       
    Wend
    
    If try_no >= 5 Then
        MsgBox "Login Error : 파일을 닫고 다시 열어주세요"
        ThisWorkbook.Close
        
        Exit Sub
    End If

        
    'emp_no = CStr(InputBox("User ID를 입력해주세요.", "User ID"))
    'emp_pass = CStr(InputBox("User Pass를 입력해주세요.", "User Pass"))
    'emp_name = Application.WorksheetFunction.VLookup(emp_no, Sheets("Config").Range("B18:C50"), 2, False)
    
                                
    Sheets("Config").cells(11, 2) = emp_no
    Sheets("Config").cells(11, 3) = emp_name

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox " " & emp_name & " 님, Welcome!"

End Sub





-------------------------------------------------------------------------------
VBA MACRO Util.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Util'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
   
   
Public Function ADD_Date(p_ref_date As Date, p_add_num As Integer, p_add_unit As String, p_holydays As Range) As Date

    Dim result_date As Date
    
    If p_add_unit = "M" Then
    
        result_date = DateAdd("m", p_add_num, p_ref_date)
        result_date = WorksheetFunction.WorkDay(result_date - 1, 1, p_holydays)
    
    ElseIf p_add_unit = "D" Then
    
        result_date = WorksheetFunction.WorkDay(p_ref_date, p_add_num, p_holydays)
    
    Else
    
        MsgBox "ADD_DATE function ERROR!!"
        
    End If
    
    ADD_Date = result_date

End Function


Public Function Find_index_from_Collection(p_collection As Collection, p_item As String) As Integer

    ' collection의 index는 1부터임...

    Dim i As Integer
    i = 1
        
    If p_collection.Count = 0 Then
        Find_index_from_Collection = -1
        Exit Function
    End If
        
    For i = 1 To p_collection.Count
    
        If p_collection(i) = p_item Then
            Find_index_from_Collection = i
            Exit Function
        End If
    
    Next
    
    Find_index_from_Collection = -1

End Function



Public Function Cnvt_to_Date(p_date As String) As Date

    Dim result_date As Date
    result_date = Left(p_date, 4) & "-" & Mid(p_date, 5, 2) & "-" & Right(p_date, 2)
    Cnvt_to_Date = result_date

End Function


Public Function File_Exists(ByVal sPathName As String, Optional Directory As Boolean) As Boolean

    'Returns True if the passed sPathName exist
    'Otherwise returns False
    
    On Error Resume Next
    
        If sPathName <> "" Then

            If IsMissing(Directory) Or Directory = False Then
                File_Exists = (Dir$(sPathName) <> "")
            Else
                File_Exists = (Dir$(sPathName, vbDirectory) <> "")
            End If
        End If
        
End Function

  
Public Function Get_Row_by_Find(p_sheet_name As String, p_col_idx As Integer, p_name_to_find As String) As Integer

    Dim i As Integer
    i = 1
    
    While i < 1000
    
        If Sheets(p_sheet_name).cells(i, p_col_idx) = p_name_to_find Then
            Get_Row_by_Find = i
            Exit Function
        End If
        
        i = i + 1
    Wend

    Get_Row_by_Find = -1

End Function
 
  
   
Public Function Get_Row_num(p_sheet_name As String, p_row_idx As Integer, p_col_idx As Integer) As Integer

    Dim i As Integer
    i = 0
    
    While Sheets(p_sheet_name).cells(p_row_idx + i, p_col_idx) <> ""
        i = i + 1
    Wend

    Get_Row_num = i

End Function
   
   
Public Sub 목록리스트_추가(Target_Cell As Range, source_name As String)

    With Target_Cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Range(source_name).Name
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

   Target_Cell.Value = Range(source_name).Offset(0, 0).Value
    
End Sub


Public Sub 이름_지정하기(new_name As String, Target_Sheet_Name As String, target_cell_name As String)

    ActiveWorkbook.Names.Add Name:=new_name, RefersTo:="=" & Target_Sheet_Name & "!" & target_cell_name

End Sub

   
Public Sub 복호화()
 
    ActiveWorkbook.Save
    ActiveWorkbook.SaveCopyAs Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - 5) & "_복호화" & ".xlsm"

End Sub


Sub SET_PRINT_AREA(end_Column)

    Dim pages As Integer
    Dim pageBegin As String
    Dim PrArea As String
    Dim i As Integer
    Dim q As Integer
    Dim nRows As Integer, nPagebreaks As Integer
    Dim R As Range
    Set R = ActiveSheet.Range(Range("A1"), Range("A7").CurrentRegion.Resize(Range("A7").CurrentRegion.Rows.Count + 1))
    
    nRows = R.Rows.Count
    
    Dim PageRows
    PageRows = 55
    If nRows > PageRows Then
      nPagebreaks = Application.WorksheetFunction.RoundUp(nRows / PageRows, 0)
      For i = 1 To nPagebreaks
         ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=R.cells(PageRows * i + 1, 1)
      Next i
    Else
      nPagebreaks = 1
      For i = 1 To nPagebreaks
         ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=R.cells(PageRows * i + 1, 1)
      Next i
    
    End If
    
    pages = ActiveSheet.HPageBreaks.Count
    pageBegin = "$A$1"
    For i = 1 To pages
      If i > 1 Then pageBegin = ActiveSheet.HPageBreaks(i - 1).Location.Address
      q = ActiveSheet.HPageBreaks(i).Location.row - 1
      PrArea = pageBegin & ":" & "$H$" & Trim$(Str$(q))
      ActiveSheet.PageSetup.PrintArea = PrArea
      ActiveSheet.PageSetup.CenterFooter = cells(q, 1)
    Next i
    
    ActiveSheet.PageSetup.PrintArea = "$A$1:" & "$" & end_Column & "$" & Trim$(Str$(q))

End Sub


Public Sub 사용자용_파일_배포()
 
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
 
    Dim file_path As String
    
    file_path = "S:\New_Swap_Platform\" & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_사용자용" & ".xlsm"
 
    ActiveWorkbook.Save
    ActiveWorkbook.SaveCopyAs file_path
    
    Workbooks.Open Filename:=file_path

    Sheets("Report관리").Select
    Range("I1") = "S:\New_Swap_Platform\Report_Files\"
    Range("K1") = "S:\New_Swap_Platform\Report_Default\"
    
    Sheets("Mailing").Select
    Range("I1") = "S:\New_Swap_Platform\Report_Files\"
    
    Sheets("Config").Select
    Range("N2") = "S:\New_Swap_Platform\FX_Files\"
    Range("N3") = "S:\New_Swap_Platform\Mug_Files\"
    
    Range("T2") = "Real_DB"
    
    Range("A1").Select
    
    Sheets("Main").Select
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close

    MsgBox "사용자용 파일 배포 완료"

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet18.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet18'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_DB_배포.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DB_배포'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DB_배포"

Public Const START_ROW_NUM As Integer = 11
Const MAX_ROW_NUM As Integer = 15000


Public Sub Sheet_DB_배포_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        If MsgBox("아래 data를 clear 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("DB_배포_기_배포_DB개수").Select
        Selection.ClearContents
        
        Range("B10").Select
              
               
    End With

    Range("C4") = "N"

    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_DB_배포_추가배포_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.cells(2, 3) & " 일자 기준으로 추가 배포대상 data를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 기존 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Call Sheet_DB_배포_Clear


        Sheets(This_Sheet_Name).Select

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("10:10").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim last_ref_date As String
        Dim sql_str As String
        Dim user_no As String
        Dim ref_seq As Integer
        

        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        '----------------------------------------------------------------------------
        ref_seq = 1

        sql_str = Range("DB_배포_MAX_REF_SEQ_SELECT").Offset(0, 0).Value2
  
        sql_str = make_SQL(sql_str, REF_DATE)
                      
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If

        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
            ref_seq = select_db_agent.DB_result_recordset(0) + 1
        End If
        
        '----------------------------------------------------------------------------
        
        
        sql_str = Range("DB_배포_추가배포_REF_LIST_SELECT_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE, ref_seq)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        

        While Not select_db_agent.DB_result_recordset.EOF
        
'    ref_date,
'    DB_SRC_1,
'    event_type,
'    DB_SRC_2,
'    client_name,
'    phone_no,
'    SEQ_NO,
'    db_input_date,
'    db_input_time,
'    age,
'    gender,
'    db_memo_1,
'    qual_mng,
'    qual_tm,
'    DUPL_CNT_1,
'    DUPL_CNT_2,
'    DUPL_LAST_DATE_1,
'    DUPL_LAST_DATE_2,
'    DUPL_LAST_DB_SRC_1


            .cells(row_idx, 2) = Cnvt_to_Date(select_db_agent.DB_result_recordset(0))
            .cells(row_idx, 3) = select_db_agent.DB_result_recordset(18)
            .cells(row_idx, 4) = select_db_agent.DB_result_recordset(6)
            .cells(row_idx, 5) = select_db_agent.DB_result_recordset(1)
            .cells(row_idx, 6) = select_db_agent.DB_result_recordset(3)
            .cells(row_idx, 7) = select_db_agent.DB_result_recordset(2)
            .cells(row_idx, 8) = select_db_agent.DB_result_recordset(4)
            .cells(row_idx, 9) = select_db_agent.DB_result_recordset(5)
            .cells(row_idx, 10) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
            .cells(row_idx, 11) = select_db_agent.DB_result_recordset(8)
            .cells(row_idx, 13) = select_db_agent.DB_result_recordset(12)
            .cells(row_idx, 14) = select_db_agent.DB_result_recordset(19)
            
            .cells(row_idx, 17) = select_db_agent.DB_result_recordset(14)
            .cells(row_idx, 18) = select_db_agent.DB_result_recordset(15)
        
            .cells(row_idx, 19) = IIf(IsNull(select_db_agent.DB_result_recordset(9)), "", select_db_agent.DB_result_recordset(9) & "  |  ") & _
                                  IIf(IsNull(select_db_agent.DB_result_recordset(10)), "", select_db_agent.DB_result_recordset(10) & "  |  ") & _
                                  select_db_agent.DB_result_recordset(11)
                                  
                                  
            .cells(row_idx, 12) = select_db_agent.DB_result_recordset(20)
            .cells(row_idx, 1) = select_db_agent.DB_result_recordset(21)
            
          
            select_db_agent.DB_result_recordset.MoveNext
        
            row_idx = row_idx + 1
        
        Wend


        ' ------------------------------------------------------------------------
        ' call log select
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) <> "Y" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                'select_db_agent.DB_result_recordset.
                
                Dim col_idx As Integer
                col_idx = 22
                
                
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                
                    .cells(row_idx, col_idx - 2) = select_db_agent.DB_result_recordset.RecordCount
                    
                    '---------------------------------------------------------
                    Rows(row_idx & ":" & row_idx).Select
                        
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                
                Else
                
                    .cells(row_idx, col_idx - 2) = ""
                
                End If
                
                          
                
                While Not select_db_agent.DB_result_recordset.EOF
            
    '담당 TM
    'Call Date
    'Call Result
    'Log
    '예약일자
    '내원여부
    
                    .cells(row_idx, col_idx) = select_db_agent.DB_result_recordset(0)
                    .cells(row_idx, col_idx + 1) = Cnvt_to_Date(select_db_agent.DB_result_recordset(1))
                    .cells(row_idx, col_idx + 2) = select_db_agent.DB_result_recordset(6)
                    .cells(row_idx, col_idx + 3) = select_db_agent.DB_result_recordset(5)
                    .cells(row_idx, col_idx + 4) = select_db_agent.DB_result_recordset(10)
                    
                    If IsNull(select_db_agent.DB_result_recordset(7)) Then
                        .cells(row_idx, col_idx + 5) = ""
                    Else
                        .cells(row_idx, col_idx + 5) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
                    End If
                    .cells(row_idx, col_idx + 6) = select_db_agent.DB_result_recordset(9)
                    
                    
                    .cells(row_idx, col_idx).Select
                    Selection.Interior.Color = 65535
                    
                    .cells(row_idx, col_idx + 2).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
              
                    select_db_agent.DB_result_recordset.MoveNext
            
                    col_idx = col_idx + 7
                            
                Wend
            
            End If
            
            
            row_idx = row_idx + 1
        

        Wend

        
        '--------------------------------------------------------------------
        ' call log TM 수 가져오기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_CNT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                
                    If select_db_agent.DB_result_recordset(0) > 0 Then
                
                       .cells(row_idx, 21) = select_db_agent.DB_result_recordset(0)
                       
                    Else
                        .cells(row_idx, 21) = ""
                        
                    End If
                        
                Else
                        .cells(row_idx, 21) = ""
                End If

            End If

            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then

                sql_str = Range("DB_배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, REF_DATE, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                       .cells(row_idx, 16) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                         & "(" & select_db_agent.DB_result_recordset(1) & ")"

                End If
                
                
                If Left(.cells(row_idx, 16), 2) = "취소" Then
                    If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                        select_db_agent.DB_result_recordset.MoveNext
                    
                        Do While Not select_db_agent.DB_result_recordset.EOF
                            If select_db_agent.DB_result_recordset(8) = "미도래" Then
                                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                    .cells(row_idx, 16) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
                                     Exit Do
                                End If
                            End If
                            select_db_agent.DB_result_recordset.MoveNext
                        Loop
                    End If
                End If

            End If

            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------



       ' ---------------------------------------------------------------------
        ' MNG 결번 표시하기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) <> "Y" Then

                sql_str = Range("구DB_작업_결번_SELECT_SQL_1").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                    .cells(row_idx, 14) = IIf(.cells(row_idx, 14) = "", "MNG_결번", .cells(row_idx, 14) & ",MNG_결번")
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend


        ' ---------------------------------------------------------------------
        ' 콜기록이 없을 경우, 최근 mng_memo 표시

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) <> "Y" Then

                If .cells(row_idx, 20) = "" Then
                
                    sql_str = Range("DB_배포_최신_MNG_MEMO_SELECT_SQL").Offset(0, 0).Value2
                    sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                    If Not select_db_agent.Select_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                
                    If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                    
                        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                        
                            .cells(row_idx, 22) = select_db_agent.DB_result_recordset(0)
                    
                            .cells(row_idx, 22).Select
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorAccent3
                                .TintAndShade = 0.599993896298105
                                .PatternTintAndShade = 0
                            End With
                                
                        End If
                    
                    End If
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend



        ' ---------------------------------------------------------------------
        ' 배포 예정과의 중복 체크

        'Dim cnt_temp As Integer
        
        'row_idx = START_ROW_NUM
        'cnt_temp = 0
        
        'While .Cells(row_idx, 2) <> ""

        '    If WorksheetFunction.CountIf(Range("I:I"), .Cells(row_idx, 9)) > 1 Then
            
        '        If .Cells(row_idx, 13) = "" Then
        '            .Cells(row_idx, 13) = "배포예정_중복"
        '        End If
            
        '    End If
            
        '    row_idx = row_idx + 1

        'Wend
        ' ---------------------------------------------------------------------


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("10:10").Select
        Selection.AutoFilter
    End If

    
    Range("C4") = "N"


    Calculate

    Range("DB_배포_현재_배포_DB개수").Select
    Selection.Copy
    Range("DB_배포_기_배포_DB개수").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B11").Select

    
    MsgBox "Query Done"


End Sub




Public Sub Sheet_DB_배포_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.cells(2, 3) & " 일자 기준으로 배포대상 data를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 기존 data는 화면에서 삭제됩니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Call Sheet_DB_배포_Clear


        Sheets(This_Sheet_Name).Select

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("10:10").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim last_ref_date As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        sql_str = Range("DB_배포_SELECT_MAX_DATE_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        last_ref_date = select_db_agent.DB_result_recordset(0)
        
        
        sql_str = Range("DB_배포_SELECT_REF_LIST_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE, last_ref_date)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        

        While Not select_db_agent.DB_result_recordset.EOF
        
'    ref_date,
'    DB_SRC_1,
'    event_type,
'    DB_SRC_2,
'    client_name,
'    phone_no,
'    SEQ_NO,
'    db_input_date,
'    db_input_time,
'    age,
'    gender,
'    db_memo_1,
'    qual_mng,
'    qual_tm,
'    DUPL_CNT_1,
'    DUPL_CNT_2,
'    DUPL_LAST_DATE_1,
'    DUPL_LAST_DATE_2,
'    DUPL_LAST_DB_SRC_1


            .cells(row_idx, 2) = Cnvt_to_Date(select_db_agent.DB_result_recordset(0))
            .cells(row_idx, 3) = select_db_agent.DB_result_recordset(18)
            .cells(row_idx, 4) = select_db_agent.DB_result_recordset(6)
            .cells(row_idx, 5) = select_db_agent.DB_result_recordset(1)
            .cells(row_idx, 6) = select_db_agent.DB_result_recordset(3)
            .cells(row_idx, 7) = select_db_agent.DB_result_recordset(2)
            .cells(row_idx, 8) = select_db_agent.DB_result_recordset(4)
            .cells(row_idx, 9) = select_db_agent.DB_result_recordset(5)
            .cells(row_idx, 10) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
            .cells(row_idx, 11) = select_db_agent.DB_result_recordset(8)
            .cells(row_idx, 13) = select_db_agent.DB_result_recordset(12)
            .cells(row_idx, 14) = select_db_agent.DB_result_recordset(19)
            
            .cells(row_idx, 17) = select_db_agent.DB_result_recordset(14)
            .cells(row_idx, 18) = select_db_agent.DB_result_recordset(15)
        
            .cells(row_idx, 19) = IIf(IsNull(select_db_agent.DB_result_recordset(9)), "", select_db_agent.DB_result_recordset(9) & "  |  ") & _
                                  IIf(IsNull(select_db_agent.DB_result_recordset(10)), "", select_db_agent.DB_result_recordset(10) & "  |  ") & _
                                  select_db_agent.DB_result_recordset(11)
                                  
          
            select_db_agent.DB_result_recordset.MoveNext
        
            row_idx = row_idx + 1
        
        Wend


        ' ------------------------------------------------------------------------
        ' call log select
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            
            'select_db_agent.DB_result_recordset.
            
            Dim col_idx As Integer
            col_idx = 22
            
            
            If select_db_agent.DB_result_recordset.RecordCount > 0 Then
            
            
                .cells(row_idx, col_idx - 2) = select_db_agent.DB_result_recordset.RecordCount
                
                '---------------------------------------------------------
                Rows(row_idx & ":" & row_idx).Select
                    
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            Else
            
                .cells(row_idx, col_idx - 2) = ""
            
            End If
            
                      
            
            While Not select_db_agent.DB_result_recordset.EOF
        
'담당 TM
'Call Date
'Call Result
'Log
'예약일자
'내원여부

                .cells(row_idx, col_idx) = select_db_agent.DB_result_recordset(0)
                .cells(row_idx, col_idx + 1) = Cnvt_to_Date(select_db_agent.DB_result_recordset(1))
                .cells(row_idx, col_idx + 2) = select_db_agent.DB_result_recordset(6)
                .cells(row_idx, col_idx + 3) = select_db_agent.DB_result_recordset(5)
                .cells(row_idx, col_idx + 4) = select_db_agent.DB_result_recordset(10)

                If IsNull(select_db_agent.DB_result_recordset(7)) Then
                    .cells(row_idx, col_idx + 5) = ""
                Else
                    .cells(row_idx, col_idx + 5) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
                End If
                .cells(row_idx, col_idx + 6) = select_db_agent.DB_result_recordset(9)
                
                
                .cells(row_idx, col_idx).Select
                Selection.Interior.Color = 65535
                
                .cells(row_idx, col_idx + 2).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                
          
                select_db_agent.DB_result_recordset.MoveNext
        
                col_idx = col_idx + 7
                        
            Wend
            
            
            row_idx = row_idx + 1
        

        Wend



        '--------------------------------------------------------------------
        ' call log TM 수 가져오기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_CNT_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            If Not select_db_agent.DB_result_recordset.EOF Then
                    
                If select_db_agent.DB_result_recordset(0) > 0 Then
            
                   .cells(row_idx, 21) = select_db_agent.DB_result_recordset(0)
                   
                Else
                    .cells(row_idx, 21) = ""
                    
                End If
            Else
                    .cells(row_idx, 21) = ""
            End If


            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------

        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            sql_str = Range("DB_배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE, .cells(row_idx, 9))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            
            If Not select_db_agent.DB_result_recordset.EOF Then
                .cells(row_idx, 16) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                         & "(" & select_db_agent.DB_result_recordset(1) & ")"
            End If
            
            
            If Left(.cells(row_idx, 16), 2) = "취소" Then
                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                    select_db_agent.DB_result_recordset.MoveNext
                
                    Do While Not select_db_agent.DB_result_recordset.EOF
                        If select_db_agent.DB_result_recordset(8) = "미도래" Then
                            If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                .cells(row_idx, 16) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
                                 Exit Do
                            End If
                        End If
                        select_db_agent.DB_result_recordset.MoveNext
                    Loop
                End If
            End If
            
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


       ' ---------------------------------------------------------------------
        ' MNG 결번 표시하기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""


            sql_str = Range("구DB_작업_결번_SELECT_SQL_1").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, .cells(row_idx, 9))
    
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
        
            If select_db_agent.DB_result_recordset.RecordCount > 0 Then
            
                .cells(row_idx, 14) = IIf(.cells(row_idx, 14) = "", "MNG_결번", .cells(row_idx, 14) & ",MNG_결번")
            
            End If
                
            
            row_idx = row_idx + 1

        Wend


        ' ---------------------------------------------------------------------
        ' 콜기록이 없을 경우, 최근 mng_memo 표시

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""


            If .cells(row_idx, 20) = "" Then
            
                sql_str = Range("DB_배포_최신_MNG_MEMO_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                    If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                    
                        .cells(row_idx, 22) = select_db_agent.DB_result_recordset(0)
                
                        .cells(row_idx, 22).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent3
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                            
                    End If
                
                End If
            
            End If

            
            row_idx = row_idx + 1

        Wend
        

        ' ---------------------------------------------------------------------
        ' 배포 예정과의 중복 체크

        Dim cnt_temp As Integer
        
        row_idx = START_ROW_NUM
        cnt_temp = 0
        
        While .cells(row_idx, 2) <> ""

            If WorksheetFunction.CountIf(Range("I:I"), .cells(row_idx, 9)) > 1 Then
            
                If .cells(row_idx, 13) = "" Then
                    .cells(row_idx, 13) = "배포예정_중복"
                End If
            
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------

        


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("10:10").Select
        Selection.AutoFilter
    End If

    
    Range("C4") = "N"


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B11").Select
    
    
    MsgBox "Query Done"


End Sub





Public Sub Sheet_DB_배포_DB_Upload()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.cells(2, 3) & " 일자 기준으로 아래 data를 DB Upload 하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        If Range("DB_배포_CHECK").Value = "False" Then
                
            MsgBox "미배정된 row가 있습니다. 다시 확인해주세요."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
            
        If .cells(4, 3) <> "Y" Then
    
            MsgBox "회수자 점검이 수행되지 않았습니다. 회수자점검 버튼을 클릭해주세요." & Chr(10) & Chr(10) & "회수자 점검 없이 upload하실려면 회수자점검 여부를 Y 로 변경해주세요."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim ref_seq As Integer
        Dim ref_seq_base As Integer
        Dim user_no As String
        Dim sql_str As String
                
                
        '------------------------------------------------------------------------
                
        Dim msg_str As String
        
        row_idx = START_ROW_NUM
        msg_str = ""
      
      
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 12) <> "배포예정" And .cells(row_idx, 12) <> "배포안함" Then
      
      
                If WorksheetFunction.CountIfs(Range("I:I"), .cells(row_idx, 9), Range("L:L"), "<>" & .cells(row_idx, 12), Range("L:L"), "<>" & "배포예정", Range("L:L"), "<>" & "배포안함") > 0 Then
      
                    msg_str = msg_str & Chr(10) & _
                                "행번호 " & row_idx & " :   " & .cells(row_idx, 5) & "   " & _
                                                            .cells(row_idx, 8) & "   " & _
                                                            .cells(row_idx, 9) & "   " & _
                                                            .cells(row_idx, 12) & "   " & _
                                                            Chr(10)
                End If
      
            End If
            
            row_idx = row_idx + 1
        
        Wend
      
        
        If msg_str <> "" Then
        
            If MsgBox("아래와 같이, 동일 전화번호에 다른 TM이 배정되었습니다." & Chr(10) & Chr(10) & _
                     "그래도 계속 DB Upload 하시겠습니까?" & Chr(10) & Chr(10) & msg_str, vbYesNo + vbDefaultButton2) = vbNo Then
        
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            
            End If
        
        End If

        
      
      
        '------------------------------------------------------------------------
      
        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        '''''''''
        ref_seq = 1
        ref_seq_base = 1
        
        '''''''''
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
                
        sql_str = ""
        
        
        
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
        '------------------------------------------------------------------------
        
        ref_seq_base = 1

        sql_str = Range("DB_배포_추가배포_MAX_REF_SEQ_SELECT").Offset(0, 0).Value2
  
        sql_str = make_SQL(sql_str, REF_DATE)
                      
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If

        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
            ref_seq_base = select_db_agent.DB_result_recordset(0) + 1
        End If
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        
        
        '---------------------------------------------------------
        ' 당일수정 관련 delete
                
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "당일수정" Then
                
                sql_str = Range("DB_배포_당일수정_SELECT_SQL").Offset(0, 0).Value2
                
'    and ref_date = :param01
'    and INPUTDB_REF_DATE = :param02
'    and INPUTDB_REF_SEQ = :param03
'    and INPUTDB_SEQ_NO = :param04
'    and DB_SRC_1 = :param05
'    and DB_SRC_2 = :param06
'    and CLIENT_NAME = :param07
'    and PHONE_NO = :param08
'    and DB_INPUT_DATE = :param09
'    and DB_INPUT_TIME = :param10
                
                sql_str = make_SQL(sql_str, _
                                    REF_DATE, _
                                    Format(.cells(row_idx, 2), "yyyyMMdd"), _
                                    .cells(row_idx, 3), _
                                    .cells(row_idx, 4), _
                                    .cells(row_idx, 5), _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 8), _
                                    .cells(row_idx, 9), _
                                    Format(.cells(row_idx, 10), "yyyyMMdd"), _
                                    .cells(row_idx, 11))
                      
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
        
                        
                If select_db_agent.DB_result_recordset.RecordCount <> 1 Then
        
                    select_db_agent.Rollback_Trans
                    Set select_db_agent = Nothing
                    MsgBox "당일수정 관련하여 Error 발생. 관리자에게 문의주세요. 행번호:" & row_idx
                    Application.Calculation = xlCalculationAutomatic
                    Application.ScreenUpdating = True
                    Exit Sub
            
                End If
                
                
                sql_str = Range("DB_배포_당일수정_DELETE_SQL").Offset(0, 0).Value2
                
'    and ref_date = :param01
'    and INPUTDB_REF_DATE = :param02
'    and INPUTDB_REF_SEQ = :param03
'    and INPUTDB_SEQ_NO = :param04
'    and DB_SRC_1 = :param05
'    and DB_SRC_2 = :param06
'    and CLIENT_NAME = :param07
'    and PHONE_NO = :param08
'    and DB_INPUT_DATE = :param09
'    and DB_INPUT_TIME = :param10
                
                sql_str = make_SQL(sql_str, _
                                    REF_DATE, _
                                    Format(.cells(row_idx, 2), "yyyyMMdd"), _
                                    .cells(row_idx, 3), _
                                    .cells(row_idx, 4), _
                                    .cells(row_idx, 5), _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 8), _
                                    .cells(row_idx, 9), _
                                    Format(.cells(row_idx, 10), "yyyyMMdd"), _
                                    .cells(row_idx, 11))
                
                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            
            
                sql_str = Range("DB_배포_당일수정_회수정보_DELETE_SQL").Offset(0, 0).Value2
                
'  REF_DATE = param01
'    and PHONE_NO = param02
'    and CLIENT_NAME = param03
'    and DB_SRC_1 = param04
'    and DB_SRC_2 = param05
                
                sql_str = make_SQL(sql_str, _
                                    REF_DATE, _
                                    .cells(row_idx, 9), _
                                    .cells(row_idx, 8), _
                                    .cells(row_idx, 5), _
                                    .cells(row_idx, 6))
                
                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            End If
        
            row_idx = row_idx + 1
        
        Wend
        
        
        '---------------------------------------------------------
        
        
        
        '---------------------------------------------------------
        ' 회수자 입력
        Dim temp_idx_1 As Integer
        Dim temp_idx_2 As Integer
        Dim 회수자이름 As String
        
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then
        
                If .cells(row_idx, 15) <> "" Then
                    
                    temp_idx_1 = 0
                    
                    Do While temp_idx_1 <= Len(.cells(row_idx, 15))
                            
                            temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 15), ",")
                    
                            If temp_idx_2 = 0 Then
                                
                                회수자이름 = Right(.cells(row_idx, 15), Len(.cells(row_idx, 15)) - temp_idx_1)
                                
                                'Exit Do
                            Else
                                회수자이름 = Mid(.cells(row_idx, 15), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 1)
                            
                            End If
                    
                            
                            sql_str = Range("DB_배포_회수자_INSERT_SQL").Offset(0, 0).Value2
    
            
    '        :param01        --        REF_DATE,
    ',        :param02        --        PHONE_NO,
    ',        (select USER_NO from user_info where USER_NAME = :param03)        --        TM_NO,
    ',        :param03        --        TM_NAME,
    ',        :param04        --        CLIENT_NAME,
    ',        :param05        --        DB_SRC_1,
    ',        :param06        --        DB_SRC_2,
    ',        :param07        --        ALT_USER_NO
            
                            sql_str = make_SQL(sql_str, _
                                                REF_DATE, _
                                                .cells(row_idx, 9), _
                                                회수자이름, _
                                                .cells(row_idx, 8), _
                                                .cells(row_idx, 5), _
                                                .cells(row_idx, 6), _
                                                user_no)
                   
                            If Not select_db_agent.Insert_update_DB(sql_str) Then
                                MsgBox "ERROR"
                            End If
                            
                            
                            If temp_idx_2 = 0 Then
                                Exit Do
                            Else
                                temp_idx_1 = temp_idx_2
                            End If
                    
                    Loop
                    
                    
                End If
                
            End If
            
            row_idx = row_idx + 1
                
        Wend
        
        
        '---------------------------------------------------------
        ' 결번 관리 입력
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then
        
                If .cells(row_idx, 14) <> "" Then
                
                    If InStr(1, cells(row_idx, 14), "결번") > 0 Then
                    
                        If cells(row_idx, 14) <> "MNG_결번" Then
                        
                            sql_str = Range("구DB_작업_결번_SELECT_SQL_2").Offset(0, 0).Value2
                        
                            sql_str = make_SQL(sql_str, _
                                            .cells(row_idx, 9), _
                                            .cells(row_idx, 8))
                
                            If Not select_db_agent.Select_DB(sql_str) Then
                                MsgBox "ERROR"
                            End If
                        
                        
                            If select_db_agent.DB_result_recordset.RecordCount = 0 Then
                            
                                sql_str = Range("구DB_작업_결번_INSERT_SQL").Offset(0, 0).Value2
            
                                sql_str = make_SQL(sql_str, _
                                            .cells(row_idx, 9), _
                                            .cells(row_idx, 8), _
                                            "DB배포 시트")
                        
                                If Not select_db_agent.Insert_update_DB(sql_str) Then
                                    MsgBox "ERROR"
                                End If
                           
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
            row_idx = row_idx + 1
                
        Wend
        
        
        '---------------------------------------------------------
        
        'sql_str = Range("DB_배포_DELETE_SQL").Offset(0, 0).Value2
        'sql_str = make_SQL(sql_str, ref_date)
        
        'If Not select_db_agent.Insert_update_DB(sql_str) Then
'       '     MsgBox "ERROR"
        'End If
        
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then
        
                If .cells(row_idx, 12) <> "배포안함" Then
                    
                    sql_str = Range("DB_배포_INSERT_SQL").Offset(0, 0).Value2
                    
            
        ': param01 --REF_DATE
        ',       :param02        -- REF_SEQ
        ',       :param03        -- DB_SRC_1
        ',       :param04        -- DB_SRC_2
        ',       :param05        -- PHONE_NO
        ',        (case when :param14 in ('배포안함') then 'N'
        '            when :param14 in ('배포예정') then 'T'
        '            else (select USER_NO from user_info where USER_NAME = :param14) end)        -- ASSIGNED_TM_NO
        ',       :param06        -- CLIENT_NAME
        ',       :param07        -- EVENT_TYPE
        ',       :param08        -- DB_INPUT_DATE
        ',       :param09        -- DB_INPUT_TIME
        ',       ''        -- AGE
        ',       ''        -- GENDER
        ',       :param10        -- DB_MEMO_1
        ',        ''     --  DB_MEMO_2
        ',        ''     --  DB_MEMO_3
        ',       (case when :param14 in ('배포안함','배포예정') then :param14 else '배포완료' end)      -- ASSIGNED_STATUS
        ',       (case when :param14 in ('배포안함','배포예정') then '' else :param14 end)             -- ASSIGNED_TM_NAME
        ',       :param11          -- INPUT_DB_QUAL
        ',       :param18          --   IS_OLD_DB
        ',       ''                --   FOLLOWING_DB_SRC
        ',       :param12          --  MNG_MEMO
        ',       :param13          --   ALT_USER_NO
        ',       :param15          -- INPUTDB_REF_DATE
        ',       :param16          -- INPUTDB_REF_SEQ
        ',       :param17          -- INPUTDB_SEQ_NO
    
                    sql_str = make_SQL(sql_str, _
                                        REF_DATE, _
                                        ref_seq_base, _
                                        .cells(row_idx, 5), _
                                        .cells(row_idx, 6), _
                                        .cells(row_idx, 9), _
                                        .cells(row_idx, 8), _
                                        .cells(row_idx, 7), _
                                        Format(.cells(row_idx, 10), "yyyyMMdd"), _
                                        .cells(row_idx, 11), _
                                        .cells(row_idx, 19), _
                                        .cells(row_idx, 13), _
                                        .cells(row_idx, 14), _
                                        user_no, _
                                        .cells(row_idx, 12), _
                                        Format(.cells(row_idx, 2), "yyyyMMdd"), _
                                        .cells(row_idx, 3), _
                                        .cells(row_idx, 4), _
                                        IIf(.cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 3) = 0, "Y", ""))
                
                    
                    If Not select_db_agent.Insert_update_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                    
                ElseIf .cells(row_idx, 12) = "배포안함" Then
                    
                    
                    ref_seq = 1
                    
    '    ref_date = :param01
    '    and db_src_2 = :param02
    '    and phone_no = :param03
    '    and assigned_tm_no = :param04
    
                    sql_str = Range("DB_배포_배포안함_MAX_REQ_SELECT").Offset(0, 0).Value2
                    
    
            
                    sql_str = make_SQL(sql_str, _
                                        REF_DATE, _
                                        .cells(row_idx, 6), _
                                        .cells(row_idx, 9), _
                                        "N")
            
                    If Not select_db_agent.Select_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
            
                    If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                        ref_seq = select_db_agent.DB_result_recordset(0) + 1
                    End If
                    
                    
                    sql_str = Range("DB_배포_INSERT_SQL").Offset(0, 0).Value2
                    
            
        ': param01 --REF_DATE
        ',       :param02        -- REF_SEQ
        ',       :param03        -- DB_SRC_1
        ',       :param04        -- DB_SRC_2
        ',       :param05        -- PHONE_NO
        ',        (case when :param14 in ('배포안함') then 'N'
        '            when :param14 in ('배포예정') then 'T'
        '            else (select USER_NO from user_info where USER_NAME = :param14) end)        -- ASSIGNED_TM_NO
        ',       :param06        -- CLIENT_NAME
        ',       :param07        -- EVENT_TYPE
        ',       :param08        -- DB_INPUT_DATE
        ',       :param09        -- DB_INPUT_TIME
        ',       ''        -- AGE
        ',       ''        -- GENDER
        ',       :param10        -- DB_MEMO_1
        ',        ''     --  DB_MEMO_2
        ',        ''     --  DB_MEMO_3
        ',       (case when :param14 in ('배포안함','배포예정') then :param14 else '배포완료' end)      -- ASSIGNED_STATUS
        ',       (case when :param14 in ('배포안함','배포예정') then '' else :param14 end)             -- ASSIGNED_TM_NAME
        ',       :param11          -- INPUT_DB_QUAL
        ',       ''                --   IS_OLD_DB
        ',       ''                --   FOLLOWING_DB_SRC
        ',       :param12          --  MNG_MEMO
        ',       :param13          --   ALT_USER_NO
        ',       :param15          -- INPUTDB_REF_DATE
        ',       :param16          -- INPUTDB_REF_SEQ
        ',       :param17          -- INPUTDB_SEQ_NO
        
                    sql_str = make_SQL(sql_str, _
                                        REF_DATE, _
                                        ref_seq, _
                                        .cells(row_idx, 5), _
                                        .cells(row_idx, 6), _
                                        .cells(row_idx, 9), _
                                        .cells(row_idx, 8), _
                                        .cells(row_idx, 7), _
                                        Format(.cells(row_idx, 10), "yyyyMMdd"), _
                                        .cells(row_idx, 11), _
                                        .cells(row_idx, 19), _
                                        .cells(row_idx, 13), _
                                        .cells(row_idx, 14), _
                                        user_no, _
                                        .cells(row_idx, 12), _
                                        Format(.cells(row_idx, 2), "yyyyMMdd"), _
                                        .cells(row_idx, 3), _
                                        .cells(row_idx, 4), _
                                        IIf(.cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 3) = 0, "Y", ""))
                
                    
                    If Not select_db_agent.Insert_update_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                    
                    
                End If
      
    
        
                '----------------------------------------------------------------
                ' input_db 테이블에 DB_QA 값 update하기
                
            End If

            row_idx = row_idx + 1
        
        Wend
                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "DB upload Done"
        
        
End Sub



Public Sub Sheet_DB_배포_회수자_점검()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("회수자 점검을 하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        '
        'End If
        
        If Range("DB_배포_CHECK").Value = "False" Then
                
            MsgBox "미배정된 row가 있습니다. 다시 확인해주세요."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Dim row_idx As Integer
        Dim tm_name_배포 As String
        Dim col_idx As Integer
        Dim i As Integer
        
        row_idx = START_ROW_NUM

        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) <> "Y" Then
        
                If .cells(row_idx, 12) <> "배포안함" And .cells(row_idx, 12) <> "배포예정" Then
                
                    tm_name_배포 = .cells(row_idx, 12)
                
                    If .cells(row_idx, 20) = "" Then
                        .cells(row_idx, 15) = ""
                    Else
                    
                        col_idx = 22
                        .cells(row_idx, 15) = ""
                    
                        For i = 1 To .cells(row_idx, 20)
                        
                            If .cells(row_idx, col_idx + 7 * (i - 1)) <> tm_name_배포 Then
                            
                                ' 기존에 회수자 이름에 없으면 추가
                                If InStr(1, .cells(row_idx, 15), .cells(row_idx, col_idx + 7 * (i - 1))) = 0 Then
                                    .cells(row_idx, 15) = .cells(row_idx, 15) & "," & .cells(row_idx, col_idx + 7 * (i - 1))
                                End If
                            End If
                            
                        Next
                        
                        If .cells(row_idx, 15) <> "" Then
                            .cells(row_idx, 15) = Right(.cells(row_idx, 15), Len(.cells(row_idx, 15)) - 1)
                        End If
                    
                    End If
                
                Else
                    .cells(row_idx, 15) = ""
                End If
                
            End If

            row_idx = row_idx + 1
            
        Wend

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("C4") = "Y"

    MsgBox "회수자 점검 Done"
        
        
End Sub




Public Sub Sheet_DB_배포_구DB_참고자료만_조희()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False


    With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim sql_str As String

        Dim REF_DATE As String

        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")

        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        ' ---------------------------------------------------------------------
        ' 회수 업무 관련
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "회수업무" Then

                sql_str = Range("DB_배포_회수업무_SELECT_SQL").Offset(0, 0).Value2
                
'        db_src_2 = :parma01
'        and client_name = :parma02
'        and phone_no = :parma03
                
                sql_str = make_SQL(sql_str, .cells(row_idx, 6), .cells(row_idx, 8), .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                .cells(row_idx, 3) = 0
                .cells(row_idx, 4) = 1
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                    
                    .cells(row_idx, 5) = select_db_agent.DB_result_recordset(0)
                    .cells(row_idx, 7) = select_db_agent.DB_result_recordset(1)
                    .cells(row_idx, 10) = select_db_agent.DB_result_recordset(2)
                    .cells(row_idx, 11) = select_db_agent.DB_result_recordset(3)
                    .cells(row_idx, 19) = select_db_agent.DB_result_recordset(4)
                    
                    
                End If
            
            
                If WorksheetFunction.CountIf(.Range("I:I"), .cells(row_idx, 9)) > 1 Then
            
                    If .cells(row_idx, 13) = "" Then
                        .cells(row_idx, 13) = "회수업무_중복"
                    End If
                    
                End If
            
            
            
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


        
        

        ' ------------------------------------------------------------------------
        ' call log select
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""


            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            
                Dim col_idx As Integer
                col_idx = 22
            
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                
                    .cells(row_idx, col_idx - 2) = select_db_agent.DB_result_recordset.RecordCount
                    
                    '---------------------------------------------------------
                    Rows(row_idx & ":" & row_idx).Select
                        
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                
                Else
                
                    .cells(row_idx, col_idx - 2) = ""
                
                End If
                
                          
                
                While Not select_db_agent.DB_result_recordset.EOF
            
    '담당 TM
    'Call Date
    'Call Result
    'Log
    '예약일자
    '내원여부
    
                    .cells(row_idx, col_idx) = select_db_agent.DB_result_recordset(0)
                    .cells(row_idx, col_idx + 1) = Cnvt_to_Date(select_db_agent.DB_result_recordset(1))
                    .cells(row_idx, col_idx + 2) = select_db_agent.DB_result_recordset(6)
                    .cells(row_idx, col_idx + 3) = select_db_agent.DB_result_recordset(5)
                    .cells(row_idx, col_idx + 4) = select_db_agent.DB_result_recordset(10)

                    If IsNull(select_db_agent.DB_result_recordset(7)) Then
                        .cells(row_idx, col_idx + 5) = ""
                    Else
                        .cells(row_idx, col_idx + 5) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
                    End If
                    .cells(row_idx, col_idx + 6) = select_db_agent.DB_result_recordset(9)
                    
                    
                    .cells(row_idx, col_idx).Select
                    Selection.Interior.Color = 65535
                    
                    .cells(row_idx, col_idx + 2).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
              
                    select_db_agent.DB_result_recordset.MoveNext
            
                    col_idx = col_idx + 7
                            
                Wend
                
            End If
                
                
            row_idx = row_idx + 1
            
    
        Wend
        
        
        '--------------------------------------------------------------------
        ' call log TM 수 가져오기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_CNT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                    If select_db_agent.DB_result_recordset(0) > 0 Then
                
                       .cells(row_idx, 21) = select_db_agent.DB_result_recordset(0)
                       
                    Else
                        .cells(row_idx, 21) = ""
                        
                    End If
                Else
                        .cells(row_idx, 21) = ""
                End If

            End If

            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, REF_DATE, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                      .cells(row_idx, 16) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                         & "(" & select_db_agent.DB_result_recordset(1) & ")"
                End If
            
            
                If Left(.cells(row_idx, 16), 2) = "취소" Then
                    If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                        select_db_agent.DB_result_recordset.MoveNext
                    
                        Do While Not select_db_agent.DB_result_recordset.EOF
                            If select_db_agent.DB_result_recordset(8) = "미도래" Then
                                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                    .cells(row_idx, 16) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
                                     Exit Do
                                End If
                            End If
                            select_db_agent.DB_result_recordset.MoveNext
                        Loop
                    End If
                End If
            
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


      ' ---------------------------------------------------------------------
        ' MNG 결번 표시하기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("구DB_작업_결번_SELECT_SQL_1").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                    .cells(row_idx, 14) = IIf(.cells(row_idx, 14) = "", "MNG_결번", .cells(row_idx, 14) & ",MNG_결번")
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend


       ' ---------------------------------------------------------------------
        ' 콜기록이 없을 경우, 최근 mng_memo 표시

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                If .cells(row_idx, 20) = "" Then
                
                    sql_str = Range("DB_배포_최신_MNG_MEMO_SELECT_SQL").Offset(0, 0).Value2
                    sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                    If Not select_db_agent.Select_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                
                    If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                    
                        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                        
                            .cells(row_idx, 22) = select_db_agent.DB_result_recordset(0)
                    
                            .cells(row_idx, 22).Select
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorAccent3
                                .TintAndShade = 0.599993896298105
                                .PatternTintAndShade = 0
                            End With
                                
                        End If
                    
                    End If
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend



    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    

    MsgBox "조회 완료"
        
        
End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet8.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet8'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet15.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet15'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_InputDB_입력.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_InputDB_입력'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "InputDB_입력"

Public Const START_ROW_NUM As Integer = 7
Const MAX_ROW_NUM As Integer = 5000


Public Sub Sheet_InputDB_입력_Clear()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        'If MsgBox(This_Sheet_Name & " sheet의 data를 clear 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("B7").Select
              
               
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub




Public Sub Sheet_InputDB_입력_DB_Upload()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.cells(2, 3) & " 일자 기준으로 아래 data를 DB Upload 하시겠습니까?" & Chr(10) & Chr(10) & _
            IIf(.cells(3, 3) = "최초/추가(회차)", "(최초/추가 입력입니다)", _
                 IIf(.cells(3, 3) = "당일재작업", "(당일 재작업입니다. 기준일자의 기존 입력 data가 삭제됩니다.)", "")), _
             vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        If .cells(5, 3) = "Y" Then
        
            MsgBox "이미 DB Upload를 진행하였습니다. " & Chr(10) & Chr(10) & _
                   "추가로 DB Upload를 진행하실려면, C5 셀을 N 으로 변경하시고 다시 시도해주세요."
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim user_no As String
        Dim sql_str As String
        Dim seq_no As Integer
        Dim ref_seq As Integer
        Dim qual_mng_result As String
                
      
        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""
        
        
        Sheets(This_Sheet_Name).Calculate
        
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        ' 당일 재작업이면 기존 data 삭제
        
        If .cells(3, 3) = "당일재작업" Then
        
            sql_str = Range("INPUT_DB_DELETE_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE)
        
            If Not select_db_agent.Insert_update_DB(sql_str) Then
'                MsgBox "ERROR"
            End If
        End If
        
        
        ref_seq = 1

        sql_str = Range("INPUT_DB_MAX_REF_SEQ_SELECT").Offset(0, 0).Value2
  
        sql_str = make_SQL(sql_str, REF_DATE)
                      
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If

        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
            ref_seq = select_db_agent.DB_result_recordset(0) + 1
        End If
        
        
        While .cells(row_idx, 2) <> ""
        
            seq_no = 1
            qual_mng_result = ""
        
            If .cells(row_idx, 1) > 0 Then
            ' 당일 중복시에 seq_no 체크
                
                
                qual_mng_result = "당일중복"
                    
                sql_str = Range("INPUT_DB_SELECT_MAX_SEQ").Offset(0, 0).Value2
            
'    ref_date = :param01
'    and db_src_1 = :param02
'    and client_name = :param03
'    and phone_no = :param04
            
                sql_str = make_SQL(sql_str, _
                                REF_DATE, _
                                .cells(row_idx, 2), _
                                .cells(row_idx, 5), _
                                .cells(row_idx, 6))
    
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
        
                If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                    seq_no = select_db_agent.DB_result_recordset(0) + 1
                End If
            
            End If
        
        
            sql_str = Range("INPUT_DB_INSERT_SQL").Offset(0, 0).Value2
            

            sql_str = make_SQL(sql_str, _
                                REF_DATE, _
                                .cells(row_idx, 2), _
                                .cells(row_idx, 5), _
                                .cells(row_idx, 6), _
                                seq_no, _
                                .cells(row_idx, 3), _
                                .cells(row_idx, 4), _
                                .cells(row_idx, 7), _
                                .cells(row_idx, 8), _
                                .cells(row_idx, 9), _
                                .cells(row_idx, 10), _
                                .cells(row_idx, 11), _
                                .cells(row_idx, 12), _
                                .cells(row_idx, 13), _
                                user_no, _
                                qual_mng_result, _
                                ref_seq, _
                                .cells(row_idx, 14), _
                                .cells(row_idx, 15))
        
            
            If Not select_db_agent.Insert_update_DB(sql_str) Then
                MsgBox "ERROR"
            End If

            row_idx = row_idx + 1
        
        Wend
                
                
        sql_str = Range("INPUT_DB_UPDATE_중복수_SQL").Offset(0, 0).Value2
            
        sql_str = make_SQL(sql_str, _
                           REF_DATE)
                           
        If Not select_db_agent.Insert_update_DB(sql_str) Then
                MsgBox "ERROR"
        End If

                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
    
        .cells(5, 3) = "Y"
    
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    
    
    
    MsgBox "DB upload 완료 ( 입력회차 : " & ref_seq & " )"
        
        
        
          
'                If .Cells(p_source_row_index, 1) <> "" Then
'                    .Cells(p_source_row_index, 1) = "이미 Booking된 Deal"
'                    select_db_agent.Rollback_Trans
'                    Exit Sub
'                End If
        

End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet9.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet14.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet14'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet4.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet4'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet7.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet7'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_InputDB_작업.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_InputDB_작업'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "InputDB_작업"
Const Target_Sheet_Name As String = "InputDB_입력"

Const START_ROW_NUM As Integer = 19
Const MAX_ROW_NUM As Integer = 5000


Public Sub Sheet_InputDB_작업_Clear()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        If MsgBox("아래 sheet data를 clear 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.Clear
       
       
       
        Range("B" & START_ROW_NUM & ":B" & MAX_ROW_NUM).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=InputDB_작업_구분자"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        Range("B" & START_ROW_NUM & ":B" & MAX_ROW_NUM).Select
    
        
       
        Range("C19") = "붙이는 곳 여기"
       
        Range("B19").Select
             
               
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Done"

End Sub


Public Sub Sheet_InputDB_작업_PreProcessing()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        If MsgBox("아래 sheet data를 다음 sheet로 복사하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Call Sheet_InputDB_입력_Clear
        
        Sheets(This_Sheet_Name).Select
        
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    
    
        'Dim total_data_cnt As Integer
        Dim row_idx As Integer
        Dim target_row_idx As Integer
                
        Dim inputdb_type As String
        
        
        'total_data_cnt = .Cells(5, 2)
        row_idx = START_ROW_NUM
        inputdb_type = ""
        
        target_row_idx = Sheet_InputDB_입력.START_ROW_NUM
        
        
        'While .Cells(row_idx, 3) <> "" Or total_data_cnt > 0
        
        While .cells(row_idx, 3) <> "" Or .cells(row_idx, 4) <> ""
        
            'If .Cells(row_idx, 3) = "" Then
            '    Rows(row_idx & ":" & row_idx).Select
            '    Selection.Delete Shift:=xlUp
            'Else
              
                'total_data_cnt = total_data_cnt - 1
                
    
                
                
                If .cells(row_idx, 2) <> "" Then
                    inputdb_type = .cells(row_idx, 2)
                Else
                    .cells(row_idx, 2) = inputdb_type
                End If
                
                
                
                
                Select Case .cells(row_idx, 2)
                
                
                
                    Case "애드인"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                         ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"   '.Cells(row_idx, 2)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 6)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 8)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 8)
                        
                        '
                        
                        Dim temp_idx_1 As Integer
                        Dim temp_idx_2 As Integer
                        Dim temp_idx_3 As Integer
                        Dim temp_idx_4 As Integer

                        temp_idx_1 = InStr(.cells(row_idx, 3), ".")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 3), ".")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 3), " ")
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 3), ":")
                        
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = IIf(IsDate(.Cells(row_idx, 3)), _
                        '                                                            Format(.Cells(row_idx, 3), "yyyyMMdd"), _
                        '                                                            Left(.Cells(row_idx, 3), 4) & Format(Mid(.Cells(row_idx, 3), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 1), "00") & Format(Mid(.Cells(row_idx, 3), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00"))
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = IIf(IsDate(.Cells(row_idx, 3)), _
                        '                                                            Format(.Cells(row_idx, 3), "hh:mm:ss"), _
                        '                                                            Format(Mid(.Cells(row_idx, 3), temp_idx_3 + 1, temp_idx_4 - temp_idx_3 - 1), "00") & ":" & Right(.Cells(row_idx, 3), 2) & ":00")
                         
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Left(.Cells(row_idx, 3), 4) & Format(Mid(.Cells(row_idx, 3), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 1), "00") & Format(Mid(.Cells(row_idx, 3), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(Mid(.Cells(row_idx, 3), temp_idx_3 + 1, temp_idx_4 - temp_idx_3 - 1), "00") & ":" & Right(.Cells(row_idx, 3), 2) & ":00"
                         
                          ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 18), "yyyyMMdd")
                         ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 18), "hh:mm:ss")
                         
                         ' 나이
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 9) = .cells(row_idx, 7)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) =
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 5) & "  |  " & .cells(row_idx, 9) & "  |  " & .cells(row_idx, 11)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                        
                        
                        Range(.cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 15).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 5).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    Case "보험임플"
                    
                        Call ProcessInsurance(.cells, row_idx, target_row_idx)
                
                    Case "방송"
                        
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = UCase(.cells(row_idx, 4))
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 4), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = IIf(Left(.cells(row_idx, 5), Len(.cells(row_idx, 5)) - 3) = "", "_", Left(.cells(row_idx, 5), Len(.cells(row_idx, 5)) - 3))
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = IIf(Mid(.Cells(row_idx, 6), 4, 1) = "-", .Cells(row_idx, 6), _
                        '                                                          "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4))
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' temp_idx_1 = InStr(.Cells(row_idx, 3), "월")
                        'temp_idx_2 = InStr(.Cells(row_idx, 3), "일")
                        
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = "2023" & Format(Left(.Cells(row_idx, 3), temp_idx_1 - 1), "00") & Format(Mid(.Cells(row_idx, 3), temp_idx_1 + 2, temp_idx_2 - temp_idx_1 - 2), "00")
                    

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = "_"
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 10) = IIf(Left(Right(.cells(row_idx, 5), 2), 1) = "남", "M", _
                                                                                IIf(Left(Right(.cells(row_idx, 5), 2), 1) = "여", "F", ""))
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =


                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "홈페이지_1", "홈페이지_2"
                    
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 4)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = .cells(row_idx, 4)

                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = UCase(.cells(row_idx, 3))
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Left(.cells(row_idx, 6), 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 9), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 9), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 13) = .cells(row_idx, 2)
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "DN_KKT_1", "DN_KKT_2"
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 5)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8) & "  |  " & .cells(row_idx, 9)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 8) & "  |  " & WorksheetFunction.Substitute(.Cells(row_idx, 8), Chr(10), "  |  ")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "모두닥_예약"
                    
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = Left(.cells(row_idx, 2), Len(.cells(row_idx, 2)) - 3)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(Left(.cells(row_idx, 2), Len(.cells(row_idx, 2)) - 3), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 4), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 4), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & "예약희망 " & .cells(row_idx, 9) & "  |  " & .cells(row_idx, 7) & "  |  " & _
                                                                              IIf(.cells(row_idx, 8) <> "", .cells(row_idx, 8) & "  |  ", "") & _
                                                                              IIf(.cells(row_idx, 10) <> "", .cells(row_idx, 10), "")

                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 10).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    Case "모두닥_전화"
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 4), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 4), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & "예약희망 " & .cells(row_idx, 9) & "  |  " & .cells(row_idx, 7) & "  |  " & _
                                                                              IIf(.cells(row_idx, 8) <> "", .cells(row_idx, 8) & "  |  ", "") & _
                                                                              IIf(.cells(row_idx, 10) <> "", .cells(row_idx, 10), "")

                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 10).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    
                    
                    
                    Case "모두닥_상담"
                    
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 4), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 4), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & .cells(row_idx, 7) & "  |  " & _
                                                                              IIf(.cells(row_idx, 8) <> "", .cells(row_idx, 8) & "  |  ", "") & _
                                                                              IIf(.cells(row_idx, 9) <> "", .cells(row_idx, 9), "")

                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                     Case "모두닥할인몰_예약"
                    
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = Left(.cells(row_idx, 2), Len(.cells(row_idx, 2)) - 3)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(Left(.cells(row_idx, 2), Len(.cells(row_idx, 2)) - 3), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 6)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 7)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 7)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 5), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 5), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & "예약희망 " & .cells(row_idx, 10) & "  |  " & _
                                                                                .cells(row_idx, 8) & "  |  " & .cells(row_idx, 4) & "  |  " & _
                                                                              IIf(.cells(row_idx, 9) <> "", .cells(row_idx, 9) & "  |  ", "") & _
                                                                              IIf(.cells(row_idx, 11) <> "", .cells(row_idx, 11), "")
                        
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 11).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                   
                     Case "모두닥할인몰_상담"
                    
                    
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = Right(.cells(row_idx, 2), 6)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(Right(.cells(row_idx, 2), 6), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 6)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 7)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 7)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 5), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 5), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & .cells(row_idx, 8) & "  |  " & .cells(row_idx, 4) & "  |  " & _
                                                                              IIf(.cells(row_idx, 9) <> "", .cells(row_idx, 9) & "  |  ", "") & _
                                                                              IIf(.cells(row_idx, 10) <> "", .cells(row_idx, 10), "")
                        
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 3).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "타불라"

                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 7)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 7)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 10), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 10), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 5) & "  |  " & .cells(row_idx, 6) & "  |  " & _
                                                                            .cells(row_idx, 11) & "  |  " & .cells(row_idx, 12) & "  |  " & .cells(row_idx, 15)
                        
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 2).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 11).Address & "," & _
                                .cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 15).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                    
                    
                    Case "페이스북_자체"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 12) & ", " & .cells(row_idx, 10)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 15)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = IIf(Left(.cells(row_idx, 16), 5) = "p:+82", _
                                                                                    "0" & Mid(.cells(row_idx, 16), 6, 2) & "-" & Mid(.cells(row_idx, 16), 8, 4) & "-" & Right(.cells(row_idx, 16), 4), _
                                                                                    Mid(.cells(row_idx, 16), 3, 3) & "-" & Mid(.cells(row_idx, 16), 6, 4) & "-" & Right(.cells(row_idx, 16), 4))
                        
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = cells(row_idx, 16)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = WorksheetFunction.Substitute(Left(.cells(row_idx, 4), 10), "-", "")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Mid(.cells(row_idx, 4), InStr(.cells(row_idx, 4), "T") + 1, 8)
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) =
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 8) & "  |  " & .Cells(row_idx, 9) & "  |  " & .Cells(row_idx, 13)
                        ' db_memo_2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 12) = .cells(row_idx, 14)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                        Range(.cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 14).Address & "," & _
                                .cells(row_idx, 15).Address & "," & _
                                .cells(row_idx, 16).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 4).Address).Select
                        
                        Selection.Interior.Color = 65535
                    
                    
                    
                    Case "지오앤플랜"
                    
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 12)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 5)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 15), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 15), "hh:mm:ss")
                        ' 나이
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 9) = .cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 10)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                        Range(.cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 15).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 10).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "빅크래프트"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = Left(.cells(row_idx, 6), 3) & "-" & Mid(.cells(row_idx, 6), 4, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 9), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 9), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 8) & "  |  " & .cells(row_idx, 10) & "  |  " & .cells(row_idx, 11)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 8) & "  |  " & WorksheetFunction.Substitute(.Cells(row_idx, 8), Chr(10), "  |  ")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 11).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "강남언니"
                        
                        'DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 9)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = Trim(.cells(row_idx, 5))
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 7)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 7)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = "20" & WorksheetFunction.Substitute(Left(.cells(row_idx, 4), 10), ". ", "")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Right(.cells(row_idx, 4), 5) & ":00"
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = .Cells(row_idx, 8)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 10) & "  |  " & .cells(row_idx, 11) & "  |  " & .cells(row_idx, 12) & "  |  " & .cells(row_idx, 13) & "  |  " & .cells(row_idx, 14)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 11).Address & "," & _
                                .cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 13).Address & "," & _
                                .cells(row_idx, 14).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "캐시닥"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = Trim(.cells(row_idx, 7))
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = Left(.cells(row_idx, 10), 3) & "-" & Mid(.cells(row_idx, 10), 4, 4) & "-" & Right(.cells(row_idx, 10), 4)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 10)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 10) = .cells(row_idx, 8)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = IIf(.cells(row_idx, 9) <> "", .cells(row_idx, 9) & "  |  ", "") & .cells(row_idx, 11) & "  |  " & .cells(row_idx, 12) & "  |  " & .cells(row_idx, 13)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 11).Address & "," & _
                                .cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 13).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                    
                    
                    Case "DS_랜딩봇", "9_DS_LD_2"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        '"0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
     
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "4K_FK"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = "4K"
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup("4K", Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 9)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 7)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        '"0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 5), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 5), "hh:mm:ss")
     
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 8)
                        '.Cells(row_idx, 4) & "  |  " & .Cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    Case "DS_35"

                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "임플란트39"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7)
                        ' db_memo_2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 12) = .cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        
                    Case "DS_400"
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "임플란트400"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7)
                        ' db_memo_2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 12) = .cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "파트너6"
                      ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7)
                        ' db_memo_2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 12) = .cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "바비톡"
                       ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 5)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 7)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = Left(.cells(row_idx, 8), 3) & "-" & Mid(.cells(row_idx, 8), 4, 4) & "-" & Right(.cells(row_idx, 8), 4)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 8)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 15), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 15), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 9) & _
                                                                                IIf(.cells(row_idx, 14) = "", "", "  |  " & .cells(row_idx, 14))
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 15).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 14).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "아이엠애드"
                      ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 7)
                        ' db_memo_2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 12) = .cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        

                    
                    
                    Case "가치브라더", "15_HID_1", "15_HID_2", "16_MT_1", "16_MT_2", "16_MT_3", "16_MT_4", "24_TK_1", "24_TK_2", "24_TK_3", "24_TK_4", "24_TK_5", "25_HID_1", "25_HID_2", "25_HID_3", "25_HID_4", "25_HID_5", "25_HID_6", "25_HID_7", "25_HID_8", "25_HID_9", "26_MT_1", "26_MT_2", "26_MT_3", "26_MT_4", "34_TK_1", "27_TB_1", "27_TB_2", "34_TK_2", "34_TK_3", "34_TK_4", "35_HID_1", "35_HID_2", "35_HID_3", "35_HID_4", "48_DB_1", "46_MT2_1", "28_DB_1", "28_DB_2"
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 7)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = IIf(Left(.cells(row_idx, 8), 2) = "82", _
                                                                                    "0" & Mid(.cells(row_idx, 8), 3, 2) & "-" & Mid(.cells(row_idx, 8), 5, 4) & "-" & Right(.cells(row_idx, 8), 4), _
                                                                                    "0" & Left(.cells(row_idx, 8), 2) & "-" & Mid(.cells(row_idx, 8), 3, 4) & "-" & Right(.cells(row_idx, 8), 4))
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 8)
                        
                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 3), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 3), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 3), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 3), ":")
                        
                        
                        
                        
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 3), 4) & _
                                                                                Format(Mid(.cells(row_idx, 3), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 3), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 3), 8) & IIf(InStr(.cells(row_idx, 3), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
                       
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = WorksheetFunction.Substitute(.cells(row_idx, 11), Chr(10), ". ") & IIf(.cells(row_idx, 12) <> "", "   |   " & .cells(row_idx, 12), "")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 12).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 11).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "카카오모먼트"

                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 5)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & IIf(.cells(row_idx, 15) = "", "", "  |  " & .cells(row_idx, 15))
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 15).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "GIOM_GDN"

                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 8)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 7)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 7)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 10), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 10), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 5) & _
                                                                                IIf(.cells(row_idx, 6) = "", "", "  |  " & .cells(row_idx, 6)) & _
                                                                                IIf(.cells(row_idx, 11) = "", "", "  |  " & .cells(row_idx, 11)) & _
                                                                                IIf(.cells(row_idx, 15) = "", "", "  |  " & .cells(row_idx, 15))
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 11).Address & "," & _
                                .cells(row_idx, 15).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
  
                    
                   Case "DS_10_INS", "DS_10_INS_2", "DS_10_INS_3", "DS_10_INS_4", "DS_10_INS_5", "DS_10_INS_6", "DS_10_INS_7", "DS_10_INS_8", "DS_10_HID", "DS_10_HID_2", "DS_10_HID_3"
                     ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 4) = "_"
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 7)
                        
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 5)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                        Selection.Interior.Color = 65535
                        
                        
                   
                   Case "DS_10_LD_1", "DS_10_LD_2", "DS_10_LD_3", "DS_10_LD_4", "DS_10_LD_5", "DS_10_LD_6", "DS_10_LD_7", "DS_TK_6", "DS_TK_7", "DS_TK_8", "DS_7_6D_1", "DS_TB_HID_2", "DS_TB_HID_3", "DS_TB_BR"
                     ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                                                
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8) & "  |  " & .cells(row_idx, 9) & "  |  " & .cells(row_idx, 10)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                        Selection.Interior.Color = 65535
                                                
                        
                    
                Case "DS_9_ins"

                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Left(.cells(row_idx, 5), 2) & "-" & Mid(.cells(row_idx, 5), 3, 4) & "-" & Right(.cells(row_idx, 5), 4)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                    
                    Case "DS_9_HID_1", "DS_9_HID_2"
                                        
                        ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 5)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                                                
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 6).Address).Select
                    
                        Selection.Interior.Color = 65535
                     
                
                  Case "DS_9_HID_3", "DS_9_HID_4"
                      ' DB src
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src name
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 5)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 6), 3, 2) & "-" & Mid(.cells(row_idx, 6), 5, 4) & "-" & Right(.cells(row_idx, 6), 4)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 9), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 9), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 9), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 9), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 9), 4) & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 9), 8) & IIf(InStr(.cells(row_idx, 9), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 3) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 4)
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                
                
                
                    
                    Case "1_DS_DN"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    Case "1_DS_DN_2"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = WorksheetFunction.Substitute(Left(.cells(row_idx, 3), 10), "-", "")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Mid(.cells(row_idx, 3), InStr(.cells(row_idx, 3), "T") + 1, 8)
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                   Case "DS_DL"
                   ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = WorksheetFunction.Substitute(Left(.cells(row_idx, 3), 10), "-", "")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Mid(.cells(row_idx, 3), InStr(.cells(row_idx, 3), "T") + 1, 8)
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                  Case "MBO"
                  ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = WorksheetFunction.Substitute(Left(.cells(row_idx, 3), 10), "-", "")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Mid(.cells(row_idx, 3), InStr(.cells(row_idx, 3), "T") + 1, 8)
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                   
                   
                   Case "DS_DL_HID"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 5), 3, 2) & "-" & Mid(.cells(row_idx, 5), 5, 4) & "-" & Right(.cells(row_idx, 5), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 8), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 8), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 8), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 8), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 8), 4) & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 8), 8) & IIf(InStr(.cells(row_idx, 8), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 9)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                   
                   Case "DS_DL_HID_M"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 5), 3, 2) & "-" & Mid(.cells(row_idx, 5), 5, 4) & "-" & Right(.cells(row_idx, 5), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 8), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 8), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 8), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 8), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 8), 4) & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 8), 8) & IIf(InStr(.cells(row_idx, 8), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 9)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                   
                   
                   Case "DS_6D_N_INS"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 5), 3, 2) & "-" & Mid(.cells(row_idx, 5), 5, 4) & "-" & Right(.cells(row_idx, 5), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 8), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 8), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 8), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 8), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 8), 4) & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 8), 8) & IIf(InStr(.cells(row_idx, 8), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 9)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                   
                   
                   
                   Case "DS_T_V_HID"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 5), 3, 2) & "-" & Mid(.cells(row_idx, 5), 5, 4) & "-" & Right(.cells(row_idx, 5), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 8), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 8), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 8), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 8), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 8), 4) & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 8), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 8), 8) & IIf(InStr(.cells(row_idx, 8), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6) & "  |  " & .cells(row_idx, 7) & "  |  " & .cells(row_idx, 9)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 9).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                     Case "DS_T_V_HID_M"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 5), 3, 2) & "-" & Mid(.cells(row_idx, 5), 5, 4) & "-" & Right(.cells(row_idx, 5), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 5)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 9), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 9), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 9), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 9), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 9), 4) & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 9), 8) & IIf(InStr(.cells(row_idx, 9), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = IIf(.cells(row_idx, 6) = "", "", .cells(row_idx, 6) & "  |  ") & _
                                                                              IIf(.cells(row_idx, 7) = "", "", .cells(row_idx, 7) & "  |  ")
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                     Case "DS_T_V_HID_M_2", "DS_T_V_HID_M_3"
                    ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 5)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 4)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "0" & Mid(.cells(row_idx, 6), 3, 2) & "-" & Mid(.cells(row_idx, 6), 5, 4) & "-" & Right(.cells(row_idx, 6), 4)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                        ' input 날짜
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = Format(.Cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = Format(.Cells(row_idx, 8), "hh:mm:ss")


                        ' input 날짜
                        temp_idx_1 = InStr(.cells(row_idx, 9), " ")
                        temp_idx_2 = InStr(temp_idx_1 + 1, cells(row_idx, 9), " ")
                        temp_idx_3 = InStr(temp_idx_2 + 1, cells(row_idx, 9), " ")
                        
                        temp_idx_4 = InStr(temp_idx_3 + 1, cells(row_idx, 9), ":")
                        

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Left(.cells(row_idx, 9), 4) & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_1 + 1, temp_idx_2 - temp_idx_1 - 2), "00") & _
                                                                                Format(Mid(.cells(row_idx, 9), temp_idx_2 + 1, temp_idx_3 - temp_idx_2 - 1), "00")

                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format( _
                                                                                        Right(.cells(row_idx, 9), 8) & IIf(InStr(.cells(row_idx, 9), "오전") > 0, "AM", "PM"), _
                                                                                        "hh:mm:ss")
     
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = IIf(.cells(row_idx, 3) = "", "", .cells(row_idx, 3) & "  |  ") & _
                                                                              IIf(.cells(row_idx, 7) = "", "", .cells(row_idx, 7) & "  |  ") & _
                                                                              IIf(.cells(row_idx, 8) = "", "", .cells(row_idx, 8))
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 9).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    
                    Case "DS_T_V"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                            
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
          
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                    
                    Case "DS_T_V_M"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = "0" & Left(.Cells(row_idx, 6), 2) & "-" & Mid(.Cells(row_idx, 6), 3, 4) & "-" & Right(.Cells(row_idx, 6), 4)
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 6)
                                                                            
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = WorksheetFunction.Substitute(Sheets(Target_Sheet_Name).Cells(target_row_idx, 6), " ", "-")
                                                                            
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                            
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
          
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 4) & "  |  " & .cells(row_idx, 8)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 8).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        
                    Case "케어랩스"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event (25.12.02 / 7->6 변경)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 6)
                        ' 이름 (25.12.02 / 8->7 변경)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 7)
                        ' 연락처 (25.12.02 / 9->8 변경)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 8)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 8)
                
                        ' input 날짜 (25.12.02 / 24->15 변경)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 15), "yyyyMMdd")
                        ' input 시간 (25.12.02 / 24->15 변경)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 15), "hh:mm:ss")
                
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                
                        ' db_memo_1 (25.12.02 / 6->5, 15->11 변경 / 16,17 삭제)
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 5) & "  |  " & _
                                                                                IIf(.cells(row_idx, 11) = "", "", .cells(row_idx, 11) & "  |  ")
                
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                
                        ' (25.12.02 / 6->5, 7->6, 8->7, 9->8, 24->15, 15->11 변경 / 16,17 삭제)
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 15).Address & "," & _
                                .cells(row_idx, 11).Address).Select
                
                        Selection.Interior.Color = 65535
                        
                        
                    Case "21_HID_1", "21_HID_2", "23_HID_1", "23_HID_2", "31_HID_1", "31_HID_2", "23_MH_TK_1", "23_MH_TK_2"
                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 4)
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 5)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = IIf(Left(.cells(row_idx, 6), 2) = "82", _
                                                                                            "0" & Mid(.cells(row_idx, 6), 3, 2) & "-" & Mid(.cells(row_idx, 6), 5, 4) & "-" & Right(.cells(row_idx, 6), 4), _
                                                                                            "0" & Mid(.cells(row_idx, 6), 1, 2) & "-" & Mid(.cells(row_idx, 6), 3, 4) & "-" & Right(.cells(row_idx, 6), 4))

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 6)
                        
                            
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 3), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 3), "hh:mm:ss")
          
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = IIf(.cells(row_idx, 7) = "", "", .cells(row_idx, 7))
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 6).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 7).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        
                        
                    Case "DS_토스"

                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "_"
                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = "_"
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = .cells(row_idx, 10)

                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 10)
                        
                            
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 8), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 8), "hh:mm:ss")
          
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 11)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 10).Address & "," & _
                                .cells(row_idx, 8).Address & "," & _
                                .cells(row_idx, 11).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        
                       
                    Case "IAI", "IAI_TK", "IAI_G2"

                        ' DB src 1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = .cells(row_idx, 2)
                        ' DB src 2
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = WorksheetFunction.VLookup(.cells(row_idx, 2), Range("CD_MAPPING_SRC_1_2"), 2, False)
                        ' Event
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = .cells(row_idx, 5)

                        ' 이름
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = .cells(row_idx, 3)
                        ' 연락처
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = IIf(Left(.cells(row_idx, 4), 2) = "82", _
                                                                                            "0" & Mid(.cells(row_idx, 4), 3, 2) & "-" & Mid(.cells(row_idx, 4), 5, 4) & "-" & Right(.cells(row_idx, 4), 4), _
                                                                                            "0" & Mid(.cells(row_idx, 4), 1, 2) & "-" & Mid(.cells(row_idx, 4), 3, 4) & "-" & Right(.cells(row_idx, 4), 4))


                        Sheets(Target_Sheet_Name).cells(target_row_idx, 14) = .cells(row_idx, 4)
                        
                            
                        ' input 날짜
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = Format(.cells(row_idx, 7), "yyyyMMdd")
                        ' input 시간
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = Format(.cells(row_idx, 7), "hh:mm:ss")
          
                        ' 나이
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 9)
                        ' 성별
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = Left(Right(.Cells(row_idx, 5), 2), 1)
                        ' db_memo_1
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = .cells(row_idx, 6)
                        ' db_memo_2
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,12) =
                        ' db_memo_3
                        'Sheets(Target_Sheet_Name).Cells(target_row_idx,13) =
                    
                    
                        Range(.cells(row_idx, 5).Address & "," & _
                                .cells(row_idx, 3).Address & "," & _
                                .cells(row_idx, 4).Address & "," & _
                                .cells(row_idx, 7).Address & "," & _
                                .cells(row_idx, 6).Address).Select
                    
                        Selection.Interior.Color = 65535
                        
                        
                    Case Else

                        MsgBox "ERROR : Unknown Input Source" & " 행번호 : " & row_idx
        
                    End Select
                    
                
                ' 번호 남기기
                Sheets(Target_Sheet_Name).cells(target_row_idx, 15) = Sheets(Target_Sheet_Name).cells(target_row_idx, 6)
    
                
                row_idx = row_idx + 1
                target_row_idx = target_row_idx + 1
                
            'End If
        
        Wend
        
        
        '-----------------------------------------------------------------------------
        
        Sheets(Target_Sheet_Name).Activate
        
        target_row_idx = Sheet_InputDB_입력.START_ROW_NUM
        
        While Sheets(Target_Sheet_Name).cells(target_row_idx, 2) <> ""
        
        
            'If WorksheetFunction.CountIf(Range("E:E"), "=" & Range("E" & target_row_idx)) > 1 Then
            
            '    Rows(target_row_idx & ":" & target_row_idx).Select
                
            '    With Selection.Interior
            '        .Pattern = xlSolid
            '        .PatternColorIndex = xlAutomatic
            '        .ThemeColor = xlThemeColorAccent2
            '        .TintAndShade = 0.799981688894314
            '        .PatternTintAndShade = 0
            '    End With
            
            'End If
            
        
        
            cells(target_row_idx, 1).Formula = "=COUNTIF(F:F,""="" & F" & target_row_idx & ")+" & "IF(E" & target_row_idx & "<>""-"",COUNTIF(E:E,""="" & E" & target_row_idx & "),1)" & "-2"
            
            cells(target_row_idx, 1).Calculate
            
            
            If cells(target_row_idx, 1) > 0 Then
            
                Rows(target_row_idx & ":" & target_row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            End If
            
            target_row_idx = target_row_idx + 1
        
        Wend
        
        
        
        
        
'        Range("C" & START_ROW_NUM & ":" & "J" & (.Cells(5, 2) + START_ROW_NUM - 1)).Select
        'Selection.ClearFormats
        
        
       ' ClearHyperlinks 하이퍼링크 지우기
       ' ClearOutline 윤곽선 지우기
       ' ClearFormat 서식 지우기
        
'        Columns("B:G").Select
'        With Selection.Font
'            .Size = 10
'        End With
'        With Selection
'            .HorizontalAlignment = xlCenter
'        End With
    
    
    
        If Not Sheets(Target_Sheet_Name).AutoFilterMode Then
            Rows("6:6").Select
            Selection.AutoFilter
        End If

        
        Range("B7").Select
               
    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Sheets(Target_Sheet_Name).cells(5, 3) = "N"
    
    Range("B7").Select
    
    MsgBox "Done"

End Sub




Public Sub Sheet_InputDB_작업_배포업무_일괄세팅하기()
    
    Dim ref_date_input As Date
    Dim REF_DATE As String
    
Try:

        On Error GoTo Catch
        ref_date_input = CDate(InputBox("배포업무를 시작하기 위해 기준일을 입력해주세요." & Chr(10) & "(ex. 2023-11-22)" & Chr(10) & Chr(10) & _
                                  "조회, InputDB_작업, InputDB_입력, DB_배포, DentWeb_연동 시트가 영향을 받습니다." & Chr(10) & Chr(10) & _
                                  "작업 중인 화면상의 데이터는 화면에서 지워집니다." & Chr(10) & Chr(10), _
                                  "배포업무 일괄세팅하기", _
                                  "", _
                            ref_date_input))
        On Error GoTo 0
                        
        GoTo Finally
        
Catch:
        MsgBox "입력된 기준일이 날짜 포맷에 맞지 않습니다. (ex. 2023-11-22)" & Chr(10) & Chr(10) & "다시 입력해주세요."
        
        Exit Sub

Finally:
        
        'Debug.Print ref_date_input
        
        REF_DATE = Format(ref_date_input, "yyyy-MM-dd")
        
        
        Sheets("조회").Select
        cells(2, 5) = REF_DATE
        cells(2, 8) = "INPUT_DB 채널별 입력시간 조회"
        Call Sheet_조회_Query
    
        Sheets("InputDB_작업").Select
        Call Sheet_InputDB_작업_Clear

        Sheets("InputDB_입력").Select
        cells(2, 3) = REF_DATE

        Sheets("DB_배포").Select
        cells(2, 3) = REF_DATE
        Call Sheet_DB_배포_Clear
        
        Sheets("DentWeb_연동").Select
        cells(2, 5) = Format(ref_date_input - 1, "yyyy-MM-dd")
        cells(2, 3) = Format(ref_date_input - 3, "yyyy-MM-dd")


        Call Sheet_DentWeb_연동_Clear
        Call Sheet_DentWeb_연동_Query
        Call Sheet_DentWeb_연동_DB_Upload

        
        Sheets("예약내원정보").Select
        cells(2, 5) = Format(ref_date_input - 1, "yyyy-MM-dd")
        cells(2, 3) = Format(ref_date_input - 3, "yyyy-MM-dd")

        Call Sheet_예약내원정보_Query
        Call Sheet_예약내원정보_DB_Update
        
        
        Sheets("InputDB_작업").Select
        
        MsgBox "배포업무 일괄 세팅 완료"
    

End Sub


Private Sub ProcessInsurance(ByRef cells As Range, ByVal row_idx As Integer, ByVal target_row_idx As Integer)
    With Sheets(Target_Sheet_Name)
        .cells(target_row_idx, 2) = cells(row_idx, 2)
        .cells(target_row_idx, 3) = cells(row_idx, 5)
        .cells(target_row_idx, 4) = "_"
        .cells(target_row_idx, 5) = cells(row_idx, 6)
        .cells(target_row_idx, 6) = cells(row_idx, 9)
        
        If cells(row_idx, 9) = "" Then
            .cells(target_row_idx, 6) = "_"
        End If
        
        .cells(target_row_idx, 7) = Format(Now, "yyyyMMdd")
        .cells(target_row_idx, 8) = Format(Now, "hh:mm:ss")
        
        .cells(target_row_idx, 9) = CalculateAgeFromYYYYMMDD(cells(row_idx, 7))
        
        .cells(target_row_idx, 11) = cells(row_idx, 4)
    End With
End Sub

Function CalculateAgeFromYYYYMMDD(birthdateStr As String) As Integer
    Dim birthdate As Date
    Dim birthYear As Integer, birthMonth As Integer, birthDay As Integer
    Dim currentDate As Date
    Dim age As Integer

    ' 문자열에서 연, 월, 일 추출
    birthYear = CInt(Left(birthdateStr, 4))
    birthMonth = CInt(Mid(birthdateStr, 5, 2))
    birthDay = CInt(Right(birthdateStr, 2))

    ' 날짜로 변환
    birthdate = DateSerial(birthYear, birthMonth, birthDay)

    currentDate = Date ' 오늘 날짜

    age = Year(currentDate) - Year(birthdate)
    ' 생일이 아직 안 지난 경우 나이에서 1 빼기
    If Month(currentDate) < Month(birthdate) Or (Month(currentDate) = Month(birthdate) And Day(currentDate) < Day(birthdate)) Then
        age = age - 1
    End If

    CalculateAgeFromYYYYMMDD = age
End Function
-------------------------------------------------------------------------------
VBA MACRO Sheet_구DB_작업.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_구DB_작업'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "구DB_작업"
Const Target_Sheet_Name As String = "DB_배포"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_구DB_작업_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        'If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = False
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_구DB_작업_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("구DB 검토 data를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 기존 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Call Sheet_구DB_작업_Clear


        Sheets(This_Sheet_Name).Select

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
                
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        sql_str = Range("구DB_작업_SELECT_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("C" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset


        '---------------------------------------------------------
        ' DB 업체명 넣기
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 3) <> ""


            If .cells(row_idx, 15) > 1 Then
                
                sql_str = Range("구DB_작업_DB업체명_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 3))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                While Not select_db_agent.DB_result_recordset.EOF
            
                    .cells(row_idx, 16) = .cells(row_idx, 16) & select_db_agent.DB_result_recordset(0) & ","
            
                    select_db_agent.DB_result_recordset.MoveNext
        
                Wend
            
            
                .cells(row_idx, 16) = Left(.cells(row_idx, 16), Len(.cells(row_idx, 16)) - 1)
                        
            
            End If
                    
            row_idx = row_idx + 1
        

        Wend


        ' ------------------------------------------------------------------------
        ' call log select
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 3) <> ""

            sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, .cells(row_idx, 3))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            
            'select_db_agent.DB_result_recordset.
            
            Dim col_idx As Integer
            col_idx = 20
            
            
            If select_db_agent.DB_result_recordset.RecordCount > 0 Then
            
            
                .cells(row_idx, col_idx - 1) = select_db_agent.DB_result_recordset.RecordCount
                
                '---------------------------------------------------------
                'Rows(row_idx & ":" & row_idx).Select
                    
                'With Selection.Interior
                '    .Pattern = xlSolid
                '    .PatternColorIndex = xlAutomatic
                '    .ThemeColor = xlThemeColorAccent1
                '    .TintAndShade = 0.799981688894314
                '    .PatternTintAndShade = 0
                'End With
            
            Else
            
                .cells(row_idx, col_idx - 1) = ""
            
            End If
            
                      
            
            While Not select_db_agent.DB_result_recordset.EOF
        
'담당 TM
'Call Date
'Call Result
'Log
'예약일자
'내원여부

                .cells(row_idx, col_idx) = select_db_agent.DB_result_recordset(0)
                .cells(row_idx, col_idx + 1) = Cnvt_to_Date(select_db_agent.DB_result_recordset(1))
                .cells(row_idx, col_idx + 2) = select_db_agent.DB_result_recordset(6)
                .cells(row_idx, col_idx + 3) = select_db_agent.DB_result_recordset(5)
                If IsNull(select_db_agent.DB_result_recordset(7)) Then
                    .cells(row_idx, col_idx + 4) = ""
                Else
                
                    If select_db_agent.DB_result_recordset(7) = "_" Then
                        .cells(row_idx, col_idx + 4) = "_"
                    Else
Try:
                        On Error GoTo Catch
                            .cells(row_idx, col_idx + 4) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
                        On Error GoTo 0
                        
                        GoTo Finally
Catch:
                        .cells(row_idx, col_idx + 4) = select_db_agent.DB_result_recordset(7)
Finally:
                        
                    End If
                
                End If
                
                .cells(row_idx, col_idx + 5) = select_db_agent.DB_result_recordset(9)
                
                
                .cells(row_idx, col_idx).Select
                Selection.Interior.Color = 65535
                
                .cells(row_idx, col_idx + 2).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                
          
                select_db_agent.DB_result_recordset.MoveNext
        
                col_idx = col_idx + 6
                        
            Wend
            
            
            row_idx = row_idx + 1
        

        Wend


        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 3) <> ""

            sql_str = Range("DB_배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, Format(Date, "yyyyMMdd"), .cells(row_idx, 3))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            
            If Not select_db_agent.DB_result_recordset.EOF Then
                .cells(row_idx, 17) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8))
                .cells(row_idx, 18) = select_db_agent.DB_result_recordset(1)
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------
        
        ' ---------------------------------------------------------------------
        ' MNG 결번 표시하기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 3) <> ""

            If .cells(row_idx, 11) = "결번" Then

                sql_str = Range("구DB_작업_결번_SELECT_SQL_1").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 3))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                    .cells(row_idx, 1) = "MNG_결번"
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend


        ' ---------------------------------------------------------------------
        ' 콜결과2가 부재 일 떄 덴트웹 결과 있으면, 콜결과2를 예약완료로 변경


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 3) <> ""

            
            If Right(.cells(row_idx, 11), 2) = "부재" And (.cells(row_idx, 17) = "이행" Or .cells(row_idx, 17) = "미이행" Or .cells(row_idx, 17) = "취소") Then
            
                .cells(row_idx, 11) = "예약완료"

            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------

    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    
    MsgBox "Query Done"



End Sub


Public Sub Sheet_구DB_작업_DB배포시트_옮기기()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("DB배포 시트로 옮기기를 하시겠습니까?" & Chr(10) & Chr(10) & "Y 표시된 행들이 DB배포 시트로 복사됩니다.", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If

        
        Dim row_idx As Integer
        Dim target_row_idx As Integer
        
        
        row_idx = START_ROW_NUM
        target_row_idx = 11
        
        While Sheets(Target_Sheet_Name).cells(target_row_idx, 2) <> ""
            target_row_idx = target_row_idx + 1
        Wend
        
        
        While .cells(row_idx, 3) <> ""
        
            If .cells(row_idx, 2) = "Y" Then
            
                Sheets(Target_Sheet_Name).cells(target_row_idx, 1) = "구DB"
                
'Input date   회차    Seq.   DB_1    DB_2    Event type  이름    연락처  In-Date In-time
                
                Sheets(Target_Sheet_Name).cells(target_row_idx, 2) = Sheets(Target_Sheet_Name).cells(2, 3)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 3) = "0"
                Sheets(Target_Sheet_Name).cells(target_row_idx, 4) = "1"
                Sheets(Target_Sheet_Name).cells(target_row_idx, 5) = "_"
                Sheets(Target_Sheet_Name).cells(target_row_idx, 6) = "구_" & .cells(row_idx, 5)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 7) = .cells(row_idx, 6)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 8) = .cells(row_idx, 4)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 9) = .cells(row_idx, 3)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 10) = .cells(row_idx, 21)
                Sheets(Target_Sheet_Name).cells(target_row_idx, 11) = "'12:00:00"
            
            
                If WorksheetFunction.CountIf(Sheets(Target_Sheet_Name).Range("I:I"), .cells(row_idx, 3)) > 1 Then
            
                    If Sheets(Target_Sheet_Name).cells(target_row_idx, 13) = "" Then
                        Sheets(Target_Sheet_Name).cells(target_row_idx, 13) = "구디비_중복"
                    End If
                    
                End If
            
                target_row_idx = target_row_idx + 1
            
            
            End If
            
            row_idx = row_idx + 1
            
        Wend

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    Sheets(Target_Sheet_Name).Select

    
    MsgBox "옮기기 완료"
    
    
End Sub




Public Sub Sheet_구DB_결번정보_DB_Upload()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(" TO_입력 으로 표시된 전화번호에 대해서 결번 정보를 DB Upload 하시겠습니까?", _
             vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Dim row_idx As Integer
        Dim sql_str As String
                
      
        row_idx = START_ROW_NUM
        sql_str = ""
        
        
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        
        While .cells(row_idx, 3) <> ""
        
        
            If .cells(row_idx, 1) = "TO_입력" And .cells(row_idx, 11) = "결번" Then
            
                sql_str = Range("구DB_작업_결번_SELECT_SQL_2").Offset(0, 0).Value2
            
                sql_str = make_SQL(sql_str, _
                                .cells(row_idx, 3), _
                                .cells(row_idx, 4))
    
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            
                If select_db_agent.DB_result_recordset.RecordCount = 0 Then
                
                    sql_str = Range("구DB_작업_결번_INSERT_SQL").Offset(0, 0).Value2

                    sql_str = make_SQL(sql_str, _
                                .cells(row_idx, 3), _
                                .cells(row_idx, 4), _
                                "")
            
                    If Not select_db_agent.Insert_update_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
               
                End If
            
            End If

            row_idx = row_idx + 1
        
        Wend
                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    
    MsgBox "결번정보 upload 완료"

End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet_SQL_실행기.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_SQL_실행기'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "SQL_실행기"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_SQL_실행기_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
    '    If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
    '
    '        Application.Calculation = xlCalculationAutomatic
    '        Application.ScreenUpdating = True
    '        Exit Sub
    '
    '    End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .cells(6, 3) = ""
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_SQL_실행기_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_SQL_실행기_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        sql_str = .cells(2, 3)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
       
        col_i = 0
       
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount

    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet12.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet12'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_DashBoard.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DashBoard'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DashBoard"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_DashBoard_Ref_Date_UP()

    Sheets(This_Sheet_Name).cells(2, 3) = Sheets(This_Sheet_Name).cells(2, 3) + 1

End Sub


Public Sub Sheet_DashBoard_Ref_Date_Down()

    Sheets(This_Sheet_Name).cells(2, 3) = Sheets(This_Sheet_Name).cells(2, 3) - 1

End Sub




Public Sub Sheet_DashBoard_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        'If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = False
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_DashBoard_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        '
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        '
        'End If


        'If Sheets(This_Sheet_Name).AutoFilterMode Then
        '    Rows("7:7").Select
        '    Selection.AutoFilter
        'End If

        Call Sheet_DashBoard_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim REF_DATE As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        '----------------------------------------------------------------------------
        .cells(row_idx, 2) = "Input DB"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_1_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        If select_db_agent.DB_result_recordset.RecordCount > 0 Then
            cells(row_idx, 2) = "합 계"
            cells(row_idx, 4).Formula = "=sum(" & "D" & (row_idx - CInt(select_db_agent.DB_result_recordset.RecordCount)) & ":D" & (row_idx - 1) & ")"
        End If
        
        row_idx = row_idx + 2
        
        '-------------------------------------------------------------------------------------------------------------------------------
        
        Dim m_ridx As Integer
        
        .cells(row_idx, 2) = "DB 배포"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        m_ridx = row_idx
        
        sql_str = Range("Dashboard_Select_2_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        If select_db_agent.DB_result_recordset.RecordCount > 0 Then
            cells(row_idx, 2) = "합 계"
            cells(row_idx, 4).Formula = "=sum(" & "D" & (row_idx - CInt(select_db_agent.DB_result_recordset.RecordCount)) & ":D" & (row_idx - 1) & ")"
        End If
            
        row_idx = row_idx + 2
        
        Call GetMissingCall(m_ridx)
        
        '-------------------------------------------------------------------------------------------------------------------------------
        
        Dim row_idx_temp As Integer
        Dim col_idx_temp As Integer
        Dim i As Integer
        
        
        .cells(row_idx, 2) = "Call Log"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_3_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        col_idx_temp = 2 + select_db_agent.DB_result_recordset.Fields.Count
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        row_idx_temp = row_idx
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        If select_db_agent.DB_result_recordset.RecordCount > 0 Then

            cells(row_idx, 2) = "합 계"
        
            For i = 4 To col_idx_temp - 1
                cells(row_idx, i).Formula = "=sum(" & Range(cells(row_idx_temp, i), cells(row_idx - 1, i)).Address & ")"
            Next i
        
            For i = row_idx_temp To row_idx - 1
                cells(i, col_idx_temp).Formula = "=" & cells(i, 4).Address & "=sum(" & Range(cells(i, 5), cells(i, col_idx_temp - 1)).Address & ")"
            Next i
        
            Range(cells(row_idx_temp, 4), cells(row_idx, col_idx_temp - 1)).Select
            Selection.Style = "Comma [0]"
        
        End If
        
        
        row_idx = row_idx + 2
        
        '-------------------------------------------------------------------------------------------------------------------------------
        
        .cells(row_idx, 2) = "중복 Call Log 점검"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_4_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "덴트웹 예약입력 누락"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_덴트웹예약누락_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "DB업체 1,2 점검"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_6_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
       '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "예약중복 log 점검"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_7_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
               
        '-------------------------------------------------------------------------------------
        .cells(row_idx, 2) = "내원여부 NULL 점검"
        
        cells(row_idx, 2).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_내원여부점검_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 2), cells(row_idx, 2 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2
        
        
        '-------------------------------------------------------------------------------------------------------------------------------

        row_idx = 8
        
        .cells(row_idx, 10) = "DentWeb 마지막 입력시간"
        
        Range(cells(row_idx, 10 - 1), cells(row_idx, 10 + 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        sql_str = Range("Dashboard_Select_5_SQL").Offset(0, 0).Value2
        
        sql_str = make_SQL(sql_str, REF_DATE)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(row_idx, 10 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 10), cells(row_idx, 10 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        row_idx = row_idx + 1
        
        Range("J" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset
        
        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        row_idx = row_idx + 2



   '     .Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
    '    .Cells(6, 5) = .Cells(2, 3) & " ~ " & .Cells(2, 5) & " 의 Call Log 전체 내역"

    End With


    'If Not Sheets(This_Sheet_Name).AutoFilterMode Then
    '    Rows("7:7").Select
    '    Selection.AutoFilter
    'End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub




Sub GetMissingCall(m_row_idx As Integer)

            Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim REF_DATE As String
        Dim sql_str As String
        Dim user_no As String
        
        Dim aSrar As String

        row_idx = m_row_idx
        
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")
        
        Set select_db_agent = New DB_Agent
        
        sql_str = "select B.ASSIGNED_TM_NO, B.ASSIGNED_TM_Name, B.CLIENT_NAME, b.phone_no From tm_call_log a, db_dist_his b where a.DB_DIST_DATE(+) = b.REF_DATE      and a.TM_NO(+) = b.ASSIGNED_TM_NO      and a.DB_SRC_2(+) = b.DB_SRC_2      and a.PHONE_NO(+) = b.PHONE_NO      and b.ref_date = '" & Format(.cells(2, 3), "yyyyMMdd") & "'      and a.ALT_DATE is null      and b.ASSIGNED_TM_NAME is not null      ;"
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        col_i = 0
        Dim inputValue As String
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
        
        inputValue = select_db_agent.DB_result_recordset.Fields(col_i).Name
                    .cells(row_idx, 5 + col_i) = Switch( _
            inputValue = "ASSIGNED_TM_NO", "TM번호", _
            inputValue = "ASSIGNED_TM_NAME", "TM이름", _
            inputValue = "CLIENT_NAME", "고객명", _
            inputValue = "PHONE_NO", "전화번호" _
        )
            col_i = col_i + 1
        Wend
        
        Range(cells(row_idx, 5), cells(row_idx, 5 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select
        
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        Range(cells(row_idx, 5), cells(row_idx, 4 + select_db_agent.DB_result_recordset.Fields.Count - 1)).Select

        row_idx = row_idx + 1

        Range("E" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)

        row_idx = row_idx + 2
        
        End With
        

End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet19.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet19'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_통계.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_통계'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "통계"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000



Public Sub Sheet_통계_조회()

    Select Case Sheets(This_Sheet_Name).cells(2, 8)
                
        Case "TM별 일일 통계"
            
            Call Sheet_통계_TM별_일일_통계_조회
            
        Case "내원환자 리스트 (내원일)"
    
            Call Sheet_통계_내원환자_리스트_내원일_조회
        
        Case "TM 기간 Performance(업체별)"
    
            Call Sheet_통계_TM별_DB업체별_Performance_조회

        Case "내원환자 리스트 (배포일)"
    
            Call Sheet_통계_내원환자_리스트_배포일_조회
            
        Case "내원환자 리스트(input date)"
    
            Call Sheet_통계_내원환자_리스트_inputdate_조회

        Case "TM 기간 Performance(업체별)_inputdate"
    
            Call Sheet_통계_TM별_DB업체별_Performance_inputdate조회

        Case Else

            MsgBox "ERROR"
        
    End Select

End Sub



Public Sub Sheet_통계_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
    '    If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
    '
    '        Application.Calculation = xlCalculationAutomatic
    '        Application.ScreenUpdating = True
    '        Exit Sub
    '
    '    End If
    
        Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Selection.Font.Bold = False
        
       ' .Cells(6, 3) = ""
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub



Public Sub Sheet_통계_내원환자_리스트_내원일_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_내원환자_리스트_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
       
'        col_i = 0
       
'        While col_i < select_db_agent.DB_result_recordset.Fields.Count
'            .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
'            col_i = col_i + 1
'        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount


        row_idx = START_ROW_NUM

        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 13) > 0 Or .cells(row_idx, 14) > 0 Then
            
                Range("B" & row_idx & ":N" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend

        
        .cells(7, 2) = "예약일"
        .cells(7, 3) = "예약시간"
        .cells(7, 4) = "차트번호"
        .cells(7, 5) = "환자이름"
        .cells(7, 6) = "전화번호"
        .cells(7, 7) = "DB정보"
        .cells(7, 8) = "DB업체"
        .cells(7, 9) = "콜기록" & Chr(10) & "기준일"
        .cells(7, 10) = "DB" & Chr(10) & "배포일"
        
        .cells(7, 11) = "TM이름"
        .cells(7, 12) = "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 13) = "타TM의" & Chr(10) & "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 14) = "차트번호" & Chr(10) & "중복수"
        

'        .Cells(7, 12) = "본인아님" & Chr(10) & "신청안함"
'        .Cells(7, 13) = "당일DB_개수"
'        .Cells(7, 14) = "예약율"
'        .Cells(7, 15) = "당일내원수"


'        Range("E7").Select
'        Range(Selection, Selection.End(xlToRight)).Select
'        Range(Selection, Selection.End(xlDown)).Select
'        ActiveWindow.SmallScroll Down:=3
'       Selection.Style = "Comma [0]"

'        Range("N7").Select
'        Range(Selection, Selection.End(xlDown)).Select
'        Selection.Style = "Percent"
'        Selection.NumberFormatLocal = "0.00%"
'        With Selection
'            .HorizontalAlignment = xlRight
'            .VerticalAlignment = xlCenter
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
        
'        Range("B7").Select
'        Range(Selection, Selection.End(xlToRight)).Select
'         With Selection
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        Selection.Font.Bold = True


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"




End Sub




Public Sub Sheet_통계_내원환자_리스트_배포일_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_내원환자_리스트_배포일_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
       
'        col_i = 0
       
'        While col_i < select_db_agent.DB_result_recordset.Fields.Count
'            .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
'            col_i = col_i + 1
'        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount


        row_idx = START_ROW_NUM

         While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 13) > 0 Or .cells(row_idx, 14) > 0 Then
            
                Range("B" & row_idx & ":N" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend

        
        .cells(7, 2) = "예약일"
        .cells(7, 3) = "예약시간"
        .cells(7, 4) = "차트번호"
        .cells(7, 5) = "환자이름"
        .cells(7, 6) = "전화번호"
        .cells(7, 7) = "DB정보"
        .cells(7, 8) = "DB업체"
        .cells(7, 9) = "콜기록" & Chr(10) & "기준일"
        .cells(7, 10) = "DB" & Chr(10) & "배포일"
        
        .cells(7, 11) = "TM이름"
        .cells(7, 12) = "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 13) = "타TM의" & Chr(10) & "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 14) = "차트번호" & Chr(10) & "중복수"


'        Range("E7").Select
'        Range(Selection, Selection.End(xlToRight)).Select
'        Range(Selection, Selection.End(xlDown)).Select
'        ActiveWindow.SmallScroll Down:=3
'       Selection.Style = "Comma [0]"

'        Range("N7").Select
'        Range(Selection, Selection.End(xlDown)).Select
'        Selection.Style = "Percent"
'        Selection.NumberFormatLocal = "0.00%"
'        With Selection
'            .HorizontalAlignment = xlRight
'            .VerticalAlignment = xlCenter
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
        
'        Range("B7").Select
'        Range(Selection, Selection.End(xlToRight)).Select
'         With Selection
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        Selection.Font.Bold = True


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"




End Sub




Public Sub Sheet_통계_TM별_일일_통계_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        If .cells(3, 3) = "일별" Then
        
            sql_str = Range("통계_일일TM통계_SELECT").Offset(0, 0).Value2
        
        ElseIf .cells(3, 3) = "기간합계" Then
        
            sql_str = Range("통계_일일TM통계_기간별_SELECT").Offset(0, 0).Value2
        
        Else
            
            sql_str = Range("통계_일일TM통계_SELECT").Offset(0, 0).Value2
        
        End If
        
        
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
       
        'col_i = 0
       
        'While col_i < select_db_agent.DB_result_recordset.Fields.Count
        '    .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
        '    col_i = col_i + 1
        'Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount


        row_idx = START_ROW_NUM

        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 4) = "합계" Then
            
                Range("B" & row_idx & ":O" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend
        
        .cells(7, 2) = "기준일"
        .cells(7, 3) = "TM번호"
        .cells(7, 4) = "TM이름"
        .cells(7, 5) = "예약" & Chr(10) & "(당일DB)"
        .cells(7, 6) = "예약" & Chr(10) & "(이전DB)"
        .cells(7, 7) = "부재"
        .cells(7, 8) = "타치과치료" & Chr(10) & "상담거부"
        .cells(7, 9) = "상담완료" & Chr(10) & "재연락"
        .cells(7, 10) = "결번"
        .cells(7, 11) = "중복"
        .cells(7, 12) = "본인아님" & Chr(10) & "신청안함"
        .cells(7, 13) = "당일DB_개수"
        .cells(7, 14) = "예약율"
        .cells(7, 15) = "당일내원수"


        Range("E7").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveWindow.SmallScroll Down:=3
        Selection.Style = "Comma [0]"

        Range("N7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Range("B7").Select
        Range(Selection, Selection.End(xlToRight)).Select
         With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Font.Bold = True


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub



Public Sub Sheet_통계_TM별_DB업체별_Performance_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_TM별DB별_SELECT_1").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        sql_str = Range("통계_TM별DB별_SELECT_2").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        sql_str = Range("통계_TM별DB별_SELECT_3").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)

        
        'col_i = 0
       
        'While col_i < select_db_agent.DB_result_recordset.Fields.Count
        '    .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
        '    col_i = col_i + 1
        'Wend



        row_idx = START_ROW_NUM

        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 2) = "합계" Or .cells(row_idx, 5) = "합계" Then
            
                Range("B" & row_idx & ":O" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend
        
        .cells(7, 2) = "DB업체1"
        .cells(7, 3) = "DB업체2"
        .cells(7, 4) = "TM번호"
        .cells(7, 5) = "TM이름"
        .cells(7, 6) = "배포DB수"
        .cells(7, 7) = "예약수"
        .cells(7, 8) = "예약율"
        .cells(7, 9) = "내원수"
        .cells(7, 10) = "내원율"
        
        
        
        Range("H7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        
        Range("J7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        Range("B7").Select
        Range(Selection, Selection.End(xlToRight)).Select
         With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Font.Bold = True

        
    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"




End Sub

Public Sub Sheet_통계_내원환자_리스트_inputdate_조회()

Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim db_input_date_1 As String
        Dim db_input_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        db_input_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        db_input_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_내원환자_리스트_inputdate_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, db_input_date_1, db_input_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
          
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = START_ROW_NUM

         While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 13) > 0 Or .cells(row_idx, 14) > 0 Then
            
                Range("B" & row_idx & ":N" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 49407
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend
        
        .cells(7, 2) = "예약일"
        .cells(7, 3) = "예약시간"
        .cells(7, 4) = "차트번호"
        .cells(7, 5) = "환자이름"
        .cells(7, 6) = "전화번호"
        .cells(7, 7) = "DB정보"
        .cells(7, 8) = "DB업체"
        .cells(7, 9) = "콜기록" & Chr(10) & "기준일"
        .cells(7, 10) = "DB" & Chr(10) & "inputdate"
        
        .cells(7, 11) = "TM이름"
        .cells(7, 12) = "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 13) = "타TM의" & Chr(10) & "동일 전화번호" & Chr(10) & "내원 중복수"
        .cells(7, 14) = "차트번호" & Chr(10) & "중복수"


    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"


End Sub

Public Sub Sheet_통계_TM별_DB업체별_Performance_inputdate조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        'If MsgBox("SQL을 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_통계_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_TM별DB별_inputdate_SELECT_1").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        sql_str = Range("통계_TM별DB별_inputdate_SELECT_2").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)
        
        
        sql_str = Range("통계_TM별DB별_inputdate_SELECT_3").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        row_idx = row_idx + CInt(select_db_agent.DB_result_recordset.RecordCount)

        
        'col_i = 0
       
        'While col_i < select_db_agent.DB_result_recordset.Fields.Count
        '    .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
        '    col_i = col_i + 1
        'Wend



        row_idx = START_ROW_NUM

        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 2) = "합계" Or .cells(row_idx, 5) = "합계" Then
            
                Range("B" & row_idx & ":O" & row_idx).Select
                
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            
            
            End If
        
        
            row_idx = row_idx + 1
        
        Wend
        
        .cells(7, 2) = "DB업체1"
        .cells(7, 3) = "DB업체2"
        .cells(7, 4) = "TM번호"
        .cells(7, 5) = "TM이름"
        .cells(7, 6) = "EVENT_TYPE"
        .cells(7, 7) = "배포DB수"
        .cells(7, 8) = "예약수"
        .cells(7, 9) = "예약율"
        .cells(7, 10) = "내원수"
        .cells(7, 11) = "내원율"
        
        With .Range("H7", .Range("H7").End(xlDown))
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        End With
        
        With .Range("H7", .Range("J7").End(xlDown))
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        End With
        
        Range("I7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        
        Range("K7").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        Range("B7").Select
        Range(Selection, Selection.End(xlToRight)).Select
         With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Font.Bold = True

        
    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"




End Sub

-------------------------------------------------------------------------------
VBA MACRO Module1.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub 매크로1()
Attribute 매크로1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로1 매크로
'

'
    Range("B12:O12").Select
  
    Range("E16").Select
End Sub
-------------------------------------------------------------------------------
VBA MACRO Module2.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub 매크로3()
Attribute 매크로3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로3 매크로
'

'

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet13.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet13'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_DB업체_결제.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DB업체_결제'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DB업체_결제"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_DB업체_결제_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        'If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
        '    Application.Calculation = xlCalculationAutomatic
        '    Application.ScreenUpdating = True
        '    Exit Sub
        
        'End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = False
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_DB업체_결제_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("DB업체_결제 data를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 기존 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Call Sheet_DB업체_결제_Clear


        Sheets(This_Sheet_Name).Select

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
                
        'user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        
        Select Case Sheets(This_Sheet_Name).cells(2, 8)
                
            Case "빅크래프트"
            
                sql_str = Range("DB업체_결제_빅크래프트_SELECT_SQL").Offset(0, 0).Value2
            
            Case "가치브라더"
        
                sql_str = Range("DB업체_결제_가치브라더_SELECT_SQL").Offset(0, 0).Value2

            Case Else

                MsgBox "ERROR"
        
        End Select
        
        
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        col_i = 0
        
        'While col_i < select_db_agent.DB_result_recordset.Fields.Count
        '    .Cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
        '    col_i = col_i + 1
        'Wend
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset



        row_idx = START_ROW_NUM
            
        While .cells(row_idx, 2) <> ""
    
    
            If .cells(2, 8) = "빅크래프트" Then
            
                If .cells(row_idx, 12) >= 10 Or _
                    .cells(row_idx, 13) >= 1 Or _
                    InStr(.cells(row_idx, 15), "무효") > 0 Or _
                    InStr(.cells(row_idx, 15), "결번") > 0 Or _
                    InStr(.cells(row_idx, 16), "결번") > 0 Or _
                    InStr(.cells(row_idx, 18), "무효") > 0 Or _
                    InStr(.cells(row_idx, 18), "결번") > 0 Or _
                    InStr(.cells(row_idx, 19), "결번") > 0 Or _
                    (.cells(row_idx, 11) > 1 And .cells(row_idx, 6) <> .cells(row_idx + 1, 6)) Then
                    
                    .cells(row_idx, 21) = "무효"
                    
                    Rows(row_idx & ":" & row_idx).Select
                    
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                                
                End If
                
                
                If .cells(row_idx, 20) >= 1 And .cells(row_idx, 21) = "" Then
                
                    .cells(row_idx, 21) = "무효의심(이름)"
                    
                    Rows(row_idx & ":" & row_idx).Select
                    
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                
                End If
                
            
            ElseIf .cells(2, 8) = "가치브라더" Then
               
                If .cells(row_idx, 12) >= 1 Or _
                    InStr(.cells(row_idx, 15), "무효") > 0 Or _
                    InStr(.cells(row_idx, 15), "결번") > 0 Or _
                    InStr(.cells(row_idx, 16), "결번") > 0 Or _
                    InStr(.cells(row_idx, 16), "본인아님") > 0 Or _
                    InStr(.cells(row_idx, 18), "무효") > 0 Or _
                    InStr(.cells(row_idx, 18), "결번") > 0 Or _
                    InStr(.cells(row_idx, 19), "결번") > 0 Or _
                    InStr(.cells(row_idx, 19), "본인아님") > 0 Or _
                    (.cells(row_idx, 11) > 1 And .cells(row_idx, 6) <> .cells(row_idx + 1, 6)) Then
                    
                    .cells(row_idx, 21) = "무효"
                    
                    Rows(row_idx & ":" & row_idx).Select
                    
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                                
                End If
               
               
                If .cells(row_idx, 20) >= 1 And .cells(row_idx, 21) = "" Then
                
                    .cells(row_idx, 21) = "무효의심(이름)"
                    
                    Rows(row_idx & ":" & row_idx).Select
                    
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                
                End If
               
            End If
            
            row_idx = row_idx + 1
    
        Wend


    End With
    


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    
    MsgBox "Query Done"



End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet_예약내원정보.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_예약내원정보'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "예약내원정보"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 10000


Public Sub Sheet_예약내원정보_Clear()

    Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("B8").Select

    MsgBox "Done"

End Sub


Public Sub Sheet_예약내원정보_Query()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    
    Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        
    With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        
        'Dim user_no As String

        row_idx = START_ROW_NUM
        col_i = 0
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 5), "yyyyMMdd")
        sql_str = ""


        '--------------------------------------------------------------------
        Set select_db_agent = New DB_Agent
        
        sql_str = Range("예약내원정보_내원여부_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount
        '.Cells(6, 5) = .Cells(2, 3) & " ~ " & .Cells(2, 5) & " 의 Call Log 전체 내역"
 
    End With
    
    

    Rows((START_ROW_NUM - 1) & ":" & (START_ROW_NUM - 1)).Select
    
    If Sheets(This_Sheet_Name).AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    Selection.AutoFilter
              
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    
    MsgBox "Query Done"

End Sub




Public Sub Sheet_예약내원정보_DB_Update()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("아래 data를 DB Update 하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    
        Dim row_idx As Integer
        Dim user_no As String
        Dim sql_str As String
                
      
        row_idx = START_ROW_NUM
        'ref_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        'ref_date_2 = Format(.Cells(2, 5), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""
        
        
                
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        
        While .cells(row_idx, 2) <> ""


            If .cells(row_idx, 4) <> "Y" And .cells(row_idx, 5) = "Y" Then
        
                sql_str = Range("예약내원정보_내원여부_UPDATE").Offset(0, 0).Value2
                
'update
'    tm_call_log a
'set
'    a.visited_YN = :param07
'    a.chart_no = :param08
'where
'    a.REF_DATE = :param01
'    and a.TM_NO = :param02
'    and a.DB_SRC_1 = :param03
'    and a.DB_SRC_2 = :param04
'    and a.PHONE_NO = :param05
'    and a.SEQ_NO = :param06    "


                sql_str = make_SQL(sql_str, _
                                    .cells(row_idx, 6), _
                                    .cells(row_idx, 7), _
                                    .cells(row_idx, 8), _
                                    .cells(row_idx, 9), _
                                    .cells(row_idx, 10), _
                                    .cells(row_idx, 11), _
                                    .cells(row_idx, 5), _
                                    IIf(.cells(row_idx, 26) = .cells(row_idx, 27), .cells(row_idx, 26), ""))

                If Not select_db_agent.Insert_update_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
            End If

            row_idx = row_idx + 1
        
        Wend
                
                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "내원여부 DB Update 완료"
        

End Sub




-------------------------------------------------------------------------------
VBA MACRO Sheet_TM_조회2.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_TM_조회2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "TM_조회_2"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 5000


Public Sub Sheet_TM_조회_2_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("B8").Select
              
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_TM_조회_2_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("콜 기록을 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_TM_조회_2_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.cells(2, 6), "yyyyMMdd")
        
        user_no = Application.WorksheetFunction.VLookup(.cells(4, 3), Sheets("Config").Range("C20:E50"), 3, False)
        
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        Select Case .cells(3, 3)
                
            Case "최초 콜 날짜"
                sql_str = Range("TM_조회_2_과거_CALL_LOG_SELECT_SQL_1").Offset(0, 0).Value2
        
            Case "최종 콜 날짜"
                sql_str = Range("TM_조회_2_과거_CALL_LOG_SELECT_SQL_2").Offset(0, 0).Value2
        
            Case "개별 콜 날짜"
                sql_str = Range("TM_조회_2_과거_CALL_LOG_SELECT_SQL_3").Offset(0, 0).Value2
        
            Case Else

                MsgBox "ERROR"
        
        End Select
        
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2, user_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(5, 3) = select_db_agent.DB_result_recordset.RecordCount
        
               ' 회수된 전번, 콜기록 회색 칠해주기
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 17) > 0 Then
            
                Rows(row_idx & ":" & row_idx).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.349986266670736
                    .PatternTintAndShade = 0
                End With
            
            End If
            
            row_idx = row_idx + 1
        
        Wend

    End With

    Columns("Q:Q").Select
    Selection.ClearContents
    Range("Q1") = 17


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub



-------------------------------------------------------------------------------
VBA MACRO Module3.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub 매크로2()
Attribute 매크로2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로2 매크로
'

'
    Range("B468:K468").Select
    With Selection.Interior

    End With
    Range("F470").Select
End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet10.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet10'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_Oracle_조회.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_Oracle_조회'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "Oracle_조회"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Long = 100000


Public Sub Sheet_Oracle_조회_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
    '    If MsgBox("아래 data를 화면에서 지우시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
    '
    '        Application.Calculation = xlCalculationAutomatic
    '        Application.ScreenUpdating = True
    '        Exit Sub
    '
    '    End If
    
        Rows((START_ROW_NUM - 1) & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .cells(6, 3) = ""
        
        Range("B8").Select
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub


Public Sub Sheet_Oracle_조회_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("Oracle 조회를 실행하시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_Oracle_조회_Clear

        

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim col_i As Integer
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        
       Select Case Sheets(This_Sheet_Name).cells(3, 3)
                
            Case "TM_CALL_LOG"
            
                sql_str = "select * from TM_CALL_LOG where ref_date >= '" & Format(.cells(2, 3), "yyyyMMdd") & "' and ref_date <= '" & Format(.cells(2, 5), "yyyyMMdd") & "' order by ref_date desc"
            
            Case "DB_DIST_HIS"
        
                sql_str = "select * from DB_DIST_HIS where ref_date >= '" & Format(.cells(2, 3), "yyyyMMdd") & "' and ref_date <= '" & Format(.cells(2, 5), "yyyyMMdd") & "' order by ref_date desc"
    
            Case "INPUT_DB"
            
                 sql_str = "select * from INPUT_DB where ref_date >= '" & Format(.cells(2, 3), "yyyyMMdd") & "' and ref_date <= '" & Format(.cells(2, 5), "yyyyMMdd") & "' order by ref_date desc"
        
            Case "DB_WITHDRAW_HIS"
            
                  sql_str = "select * from DB_WITHDRAW_HIS where ref_date >= '" & Format(.cells(2, 3), "yyyyMMdd") & "' and ref_date <= '" & Format(.cells(2, 5), "yyyyMMdd") & "' order by ref_date desc"
            
            Case "MNG_결번_관리"
    
                   sql_str = "select * from MNG_결번_관리 where alt_date >= '" & Format(.cells(2, 3), "yyyyMMdd") & "' and alt_date <= '" & Format(.cells(2, 5), "yyyyMMdd") & "' order by alt_date desc"
            
            Case "TM_CALL_LOG(중복제거)+DENTWEB"
    
                    Call Sheet_TM_CALL_LOG_DENTWEB_조회
                    Exit Sub
                    
    
            Case Else
    
                MsgBox "ERROR"
            
        End Select
        
        
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
       
        col_i = 0
       
        While col_i < select_db_agent.DB_result_recordset.Fields.Count
            .cells(7, 2 + col_i) = select_db_agent.DB_result_recordset.Fields(col_i).Name
            col_i = col_i + 1
        Wend
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .cells(6, 3) = select_db_agent.DB_result_recordset.RecordCount

    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub


Public Sub Sheet_TM_CALL_LOG_DENTWEB_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    On Error GoTo CLEAN_FAIL

    With Sheets(This_Sheet_Name)


        If .AutoFilterMode Then
            .Rows("7:7").AutoFilter
        End If


        Call Sheet_Oracle_조회_Clear

        Dim select_db_agent As DB_Agent
        Dim sql_raw As String, sql_str As String
        Dim param01 As String, param02 As String
        Dim col_i As Long


        param01 = Format(.cells(2, 3).Value, "yyyymmdd") ' C2
        param02 = Format(.cells(2, 5).Value, "yyyymmdd") ' E2


        sql_raw = Sheets("SQL").Range("C81").Value2

        ' 바인드 치환 (프로젝트에 make_SQL 함수가 이미 있다면 그걸 사용)
        ' 우선 make_SQL이 있다고 가정:
        On Error Resume Next
        sql_str = make_SQL(sql_raw, param01, param02)
        If Err.Number <> 0 Or Len(sql_str) = 0 Then
            ' make_SQL이 없거나 실패하면 Replace로 직접 치환
            Err.Clear
            sql_str = Replace(sql_raw, ":param01", "'" & param01 & "'")
            sql_str = Replace(sql_str, ":param02", "'" & param02 & "'")
        End If
        On Error GoTo CLEAN_FAIL

        ' DB 실행
        Set select_db_agent = New DB_Agent
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "DB 조회 중 오류가 발생했습니다.", vbCritical
            GoTo CLEAN_OK
        End If

        ' 헤더 출력
        For col_i = 0 To select_db_agent.DB_result_recordset.Fields.Count - 1
            .cells(7, 2 + col_i).Value = select_db_agent.DB_result_recordset.Fields(col_i).Name
        Next col_i

        ' 데이터 출력
        .Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        ' 건수
        .cells(6, 3).Value = select_db_agent.DB_result_recordset.RecordCount

        ' 자동필터 켜기
        If Not .AutoFilterMode Then
            .Rows("7:7").AutoFilter
        End If

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Range("B8").Select
    MsgBox "조회 완료"
    Exit Sub

CLEAN_FAIL:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical
    Exit Sub

CLEAN_OK:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet17.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet17'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_DB_검색.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_DB_검색'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "DB_검색"

Public Const START_ROW_NUM As Integer = 11
Const MAX_ROW_NUM As Integer = 5000


Public Sub Sheet_DB_검색_Clear()

    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
    
        Sheets(This_Sheet_Name).Select
   
    
        If MsgBox("아래 data를 clear 하시겠습니까??", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
        Rows(START_ROW_NUM & ":" & MAX_ROW_NUM).Select
        Selection.ClearContents
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Range("DB_검색_기_배포_DB개수").Select
        Selection.ClearContents
        
        Range("B10").Select
              
               
    End With

    Range("C4") = "N"

    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub








Public Sub Sheet_DB_검색_구DB_참고자료만_조희()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False


    With Sheets(This_Sheet_Name)

        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim sql_str As String

        Dim REF_DATE As String

        row_idx = START_ROW_NUM
        REF_DATE = Format(.cells(2, 3), "yyyyMMdd")

        sql_str = ""

     
        Set select_db_agent = New DB_Agent
        
        
        ' ---------------------------------------------------------------------
        ' 회수 업무 관련
        
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "회수업무" Then

                sql_str = Range("DB_배포_회수업무_SELECT_SQL").Offset(0, 0).Value2
                
'        db_src_2 = :parma01
'        and client_name = :parma02
'        and phone_no = :parma03
                
                sql_str = make_SQL(sql_str, .cells(row_idx, 6), .cells(row_idx, 8), .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                .cells(row_idx, 3) = 0
                .cells(row_idx, 4) = 1
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                    
                    .cells(row_idx, 5) = select_db_agent.DB_result_recordset(0)
                    .cells(row_idx, 7) = select_db_agent.DB_result_recordset(1)
                    .cells(row_idx, 10) = select_db_agent.DB_result_recordset(2)
                    .cells(row_idx, 11) = select_db_agent.DB_result_recordset(3)
                    .cells(row_idx, 19) = select_db_agent.DB_result_recordset(4)
                    
                    
                End If
            
            
                If WorksheetFunction.CountIf(.Range("I:I"), .cells(row_idx, 9)) > 1 Then
            
                    If .cells(row_idx, 13) = "" Then
                        .cells(row_idx, 13) = "회수업무_중복"
                    End If
                    
                End If
            
            
            
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


        
        

        ' ------------------------------------------------------------------------
        ' call log select
        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""


            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
            
                Dim col_idx As Integer
                col_idx = 22
            
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                
                    .cells(row_idx, col_idx - 2) = select_db_agent.DB_result_recordset.RecordCount
                    
                    '---------------------------------------------------------
                    Rows(row_idx & ":" & row_idx).Select
                        
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                
                Else
                
                    .cells(row_idx, col_idx - 2) = ""
                
                End If
                
                          
                
                While Not select_db_agent.DB_result_recordset.EOF
            
    '담당 TM
    'Call Date
    'Call Result
    'Log
    '예약일자
    '내원여부
    
                    .cells(row_idx, col_idx) = select_db_agent.DB_result_recordset(0)
                    .cells(row_idx, col_idx + 1) = Cnvt_to_Date(select_db_agent.DB_result_recordset(1))
                    .cells(row_idx, col_idx + 2) = select_db_agent.DB_result_recordset(6)
                    .cells(row_idx, col_idx + 3) = select_db_agent.DB_result_recordset(5)
                    .cells(row_idx, col_idx + 4) = select_db_agent.DB_result_recordset(10)

                    If IsNull(select_db_agent.DB_result_recordset(7)) Then
                        .cells(row_idx, col_idx + 5) = ""
                    Else
                        .cells(row_idx, col_idx + 5) = Cnvt_to_Date(select_db_agent.DB_result_recordset(7))
                    End If
                    .cells(row_idx, col_idx + 6) = select_db_agent.DB_result_recordset(9)
                    
                    
                    .cells(row_idx, col_idx).Select
                    Selection.Interior.Color = 65535
                    
                    .cells(row_idx, col_idx + 2).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
              
                    select_db_agent.DB_result_recordset.MoveNext
            
                    col_idx = col_idx + 7
                            
                Wend
                
            End If
                
                
            row_idx = row_idx + 1
            
    
        Wend
        
        
        '--------------------------------------------------------------------
        ' call log TM 수 가져오기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_SELECT_TM_CALL_LOG_CNT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                    If select_db_agent.DB_result_recordset(0) > 0 Then
                
                       .cells(row_idx, 21) = select_db_agent.DB_result_recordset(0)
                       
                    Else
                        .cells(row_idx, 21) = ""
                        
                    End If
                Else
                        .cells(row_idx, 21) = ""
                End If

            End If

            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기


        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""
        
            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("DB_배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, REF_DATE, .cells(row_idx, 9))
            
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                
                If Not select_db_agent.DB_result_recordset.EOF Then
                      .cells(row_idx, 16) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                         & "(" & select_db_agent.DB_result_recordset(1) & ")"
                End If
            
            
                If Left(.cells(row_idx, 16), 2) = "취소" Then
                    If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                        select_db_agent.DB_result_recordset.MoveNext
                    
                        Do While Not select_db_agent.DB_result_recordset.EOF
                            If select_db_agent.DB_result_recordset(8) = "미도래" Then
                                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                    .cells(row_idx, 16) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
                                     Exit Do
                                End If
                            End If
                            select_db_agent.DB_result_recordset.MoveNext
                        Loop
                    End If
                End If
            
            End If
            
            row_idx = row_idx + 1

        Wend
        ' ---------------------------------------------------------------------


      ' ---------------------------------------------------------------------
        ' MNG 결번 표시하기

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                sql_str = Range("구DB_작업_결번_SELECT_SQL_1").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, .cells(row_idx, 9))
        
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
            
                If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                
                    .cells(row_idx, 14) = IIf(.cells(row_idx, 14) = "", "MNG_결번", .cells(row_idx, 14) & ",MNG_결번")
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend


       ' ---------------------------------------------------------------------
        ' 콜기록이 없을 경우, 최근 mng_memo 표시

        row_idx = START_ROW_NUM
        
        While .cells(row_idx, 2) <> ""

            If .cells(row_idx, 1) = "구DB" Or .cells(row_idx, 1) = "회수업무" Or .cells(row_idx, 1) = "당일수정" Then

                If .cells(row_idx, 20) = "" Then
                
                    sql_str = Range("DB_배포_최신_MNG_MEMO_SELECT_SQL").Offset(0, 0).Value2
                    sql_str = make_SQL(sql_str, .cells(row_idx, 9))
            
                    If Not select_db_agent.Select_DB(sql_str) Then
                        MsgBox "ERROR"
                    End If
                
                    If select_db_agent.DB_result_recordset.RecordCount > 0 Then
                    
                        If Not IsNull(select_db_agent.DB_result_recordset(0)) Then
                        
                            .cells(row_idx, 22) = select_db_agent.DB_result_recordset(0)
                    
                            .cells(row_idx, 22).Select
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorAccent3
                                .TintAndShade = 0.599993896298105
                                .PatternTintAndShade = 0
                            End With
                                
                        End If
                    
                    End If
                
                End If
                
            End If
            
            row_idx = row_idx + 1

        Wend



    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    

    MsgBox "조회 완료"
        
        
End Sub




