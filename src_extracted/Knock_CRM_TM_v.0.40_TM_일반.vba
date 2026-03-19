olevba 0.60.2 on Python 3.11.7 - http://decalage.info/python/oletools
===============================================================================
FILE: Knock_CRM_TM_v.0.40_TM_일반.xlsm
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
VBA MACRO Sheet9.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_Config.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_Config'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_예약조회.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_예약조회'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "예약조회"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 1000


Public Sub Sheet_예약관리_Clear()

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


Public Sub Sheet_예약조회_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("예약 정보를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_예약관리_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim reserv_date_1 As String
        'Dim reserv_date_2 As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        reserv_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        'reserv_date_2 = Format(.Cells(2, 6), "yyyyMMdd")
        
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        Select Case .Cells(3, 3)
                
            Case "취소 포함"
                sql_str = Range("예약관리_예약_조회_취소포함_SELECT").Offset(0, 0).Value2
        
            Case "취소 제외"
                sql_str = Range("예약관리_예약_조회_취소제외_SELECT").Offset(0, 0).Value2
        
            Case Else
                MsgBox "ERROR"
        
        End Select
        
        'sql_str = make_SQL(sql_str, IIf(reserv_date_1 = "", "ALL", reserv_date_1), IIf(reserv_date_2 = "", "ALL", reserv_date_2), user_no)
        sql_str = make_SQL(sql_str, reserv_date_1, user_no)
        
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(5, 3) = select_db_agent.DB_result_recordset.RecordCount

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


Public Sub Sheet_예약조회_당일입력_예약조회_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("당일입력 예약 정보를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If

        Call Sheet_예약관리_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim call_date_1 As String
        'Dim reserv_date_2 As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        call_date_1 = Format(Date, "yyyyMMdd")
        'reserv_date_2 = Format(.Cells(2, 6), "yyyyMMdd")
        
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
  
        sql_str = Range("예약관리_당일입력_예약_조회_SELECT").Offset(0, 0).Value2
        
        'sql_str = make_SQL(sql_str, IIf(reserv_date_1 = "", "ALL", reserv_date_1), IIf(reserv_date_2 = "", "ALL", reserv_date_2), user_no)
        sql_str = make_SQL(sql_str, call_date_1, user_no)
        
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        '.Cells(5, 3) = select_db_agent.DB_result_recordset.RecordCount

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


Public Sub Sheet_예약관리_행추가()

'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'
'    With Sheets(This_Sheet_Name)
'
'        Dim row_idx As Integer
'
'        row_idx = ActiveCell.Row
'
'        Sheets(This_Sheet_Name).Select
'        Rows(row_idx & ":" & row_idx).Select
'        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'        With Selection.Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorAccent3
'            .TintAndShade = 0.599993896298105
'            .PatternTintAndShade = 0
'        End With
'
'        .Cells(row_idx, 2) = Date
'        .Cells(row_idx, 4) = Range("Cur_User_No").Offset(0, 1).Value2
'        .Cells(row_idx, 5) = .Cells(row_idx + 1, 5)
'        .Cells(row_idx, 6) = .Cells(row_idx + 1, 6)
'        .Cells(row_idx, 7) = .Cells(row_idx + 1, 7)
'        .Cells(row_idx, 8) = .Cells(row_idx + 1, 8)
'
'        Cells(row_idx, 9).Select
'
'    End With

'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
    

End Sub




Public Sub Sheet_예약관리_CRM_입력_복사_개별()

'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False

'    With Sheets(This_Sheet_Name)
        
'        Dim row_idx As Integer
'
'        row_idx = ActiveCell.Row
'
'        Sheets(This_Sheet_Name).Select
'
'
'        If .Cells(row_idx, 2) <> Date Then
'            MsgBox "당일 입력 call log가 아닙니다."
'            Application.Calculation = xlCalculationAutomatic
'            Application.ScreenUpdating = True
'            Exit Sub
'        End If
'
'        If .Cells(row_idx, 10) = "" Then
'            MsgBox "콜 결과값이 없습니다."
'            Application.Calculation = xlCalculationAutomatic
'            Application.ScreenUpdating = True
'            Exit Sub
'        End If
'
'        '--------------------------------------------------------------------------------------
'        Dim target_row_idx As Integer
'
'        target_row_idx = 7
'
'        While Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) <> ""
'            target_row_idx = target_row_idx + 1
'        Wend
'        '--------------------------------------------------------------------------------------
'
'' 당일/직전 Calldate
''담당자
''DB업체
''DB정보
''이름
''연락처
''상담내용
''결과
''예약일
''예약시간
''내원여부
''담당자 메모
'
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) = .Cells(row_idx, 2)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 3) = .Cells(row_idx, 4)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 4) = .Cells(row_idx, 5)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 5) = .Cells(row_idx, 6)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 7)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = .Cells(row_idx, 8)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = .Cells(row_idx, 9)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 10)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = .Cells(row_idx, 11)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 12)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 13)
'        Sheets(Target_Sheet_Name).Cells(target_row_idx, 13) = .Cells(row_idx, 14)
'
'        'Cells(row_idx, 9).Select
'
'    End With
'
'    Application.Calculation = xlCalculationAutomatic
'    Application.ScreenUpdating = True
'
'    MsgBox "복사 Done"

End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet_CRM_입력.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_CRM_입력'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "CRM_입력"

Public Const START_ROW_NUM As Integer = 7
Const MAX_ROW_NUM As Integer = 1000


Public Sub Sheet_CRM_입력_Clear()

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
        
        
        Range("C5") = "N"
        
        Range("B7").Select
              
               
    End With


    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    
    'MsgBox "Done"

End Sub



Public Sub Sheet_CRM_입력_CRM_Upload()


    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox(.Cells(2, 3) & " 일자 기준으로 아래 화면의 data를 CRM Upload 하시겠습니까?" & Chr(10) & Chr(10) & "(기준일자의 기존 콜 기록이 있을 경우, 추가로 Upload 됩니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
        
        
        If .Cells(5, 3) = "Y" Then
        
            If MsgBox(.Cells(2, 3) & " 일자 기준으로 CRM Upload 를 하신 적이 있습니다." & Chr(10) & Chr(10) & "그래도 계속 진행하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            
            End If
        
        End If
    
    
        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim user_no As String
        Dim sql_str As String
        Dim seq_no As Integer
                
      
        row_idx = START_ROW_NUM
        REF_DATE = Format(.Cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""
        
        
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
        
        'sql_str = Range("INPUT_DB_DELETE_SQL").Offset(0, 0).Value2
        'sql_str = make_SQL(sql_str, REF_DATE)
        
        'If Not select_db_agent.Insert_update_DB(sql_str) Then
'            MsgBox "ERROR"
        'End If
        
        
        While .Cells(row_idx, 2) <> ""
        
            sql_str = Range("CRM_입력_MAX_SEQ_SELECT_SQL").Offset(0, 0).Value2
            
'    ref_date = :param01
'    and tm_no = (select user_no from user_info where user_name = :param02)
'    and db_src_2 = :param03
'    and phone_no = :param04
            
            sql_str = make_SQL(sql_str, _
                                REF_DATE, _
                                .Cells(row_idx, 3), _
                                .Cells(row_idx, 4), _
                                .Cells(row_idx, 7))
    
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
    
            If IsNull(select_db_agent.DB_result_recordset(0)) Then
                seq_no = 1
            Else
                seq_no = select_db_agent.DB_result_recordset(0) + 1
            End If
        
        
            Dim db_dist_date As String
            
            db_dist_date = ""
            
            If .Cells(row_idx, 2) = "당일" Then
                
                db_dist_date = REF_DATE
            
            ElseIf .Cells(row_idx, 2) = Date Then

'        tm_no = (select user_no from user_info where user_name = :param01)
'        and db_src_2 = :param02
'        and phone_no = :param03

                sql_str = Range("CRM_입력_최근_DB_DIST_DATE_SELECT_SQL").Offset(0, 0).Value2
                sql_str = make_SQL(sql_str, _
                                    .Cells(row_idx, 3), _
                                    .Cells(row_idx, 4), _
                                    .Cells(row_idx, 7))
    
                If Not select_db_agent.Select_DB(sql_str) Then
                    MsgBox "ERROR"
                End If
                
                If select_db_agent.DB_result_recordset.EOF Then
                    db_dist_date = "_"
                Else
                    db_dist_date = select_db_agent.DB_result_recordset(0)
                End If
                
            Else
                MsgBox "ERROR"
            End If
        
       
        
        
            sql_str = Range("CRM_입력_INSERT_SQL").Offset(0, 0).Value2
            

'   : param01   --REF_DATE
'   : param02   --tm_name
',   :param03    -- DB_SRC_2
',   :param04    --PHONE_NO
',   :param05    -- SEQ_NO
',   :param06    -- EVENT_TYPE
',   :param07    -- CLIENT_NAME
',   :param02    -- TM_NAME
',   :param08    -- call_MEMO_1
',   :param09    -- call_MEMO_2
',   :param10    -- CALL_RESULT
',   :param11    -- RESERV_DATE
',   :param12    -- RESERV_TIME
',   :param13    -- VISITED_YN
',   :param14    -- tm_MEMO
',   :param15    -- DB_DIST_DATE
',   :param16    -- IN_OUT_TYPE
',   :param17    -- ALT_USER_NO


            sql_str = make_SQL(sql_str, _
                                REF_DATE, _
                                .Cells(row_idx, 3), _
                                .Cells(row_idx, 4), _
                                .Cells(row_idx, 7), _
                                seq_no, _
                                .Cells(row_idx, 5), _
                                .Cells(row_idx, 6), _
                                WorksheetFunction.Substitute(.Cells(row_idx, 8), "'", " "), _
                                "", _
                                .Cells(row_idx, 9), _
                                Format(.Cells(row_idx, 10), "yyyyMMdd"), _
                                Format(.Cells(row_idx, 11), "hh:mm:ss"), _
                                .Cells(row_idx, 12), _
                                WorksheetFunction.Substitute(.Cells(row_idx, 13), "'", " "), _
                                db_dist_date, _
                                "O", _
                                user_no)
        
            
            If Not select_db_agent.Insert_update_DB(sql_str) Then
                MsgBox "ERROR"
            End If

            row_idx = row_idx + 1
        
        Wend
                

                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
    
        .Cells(5, 3) = "Y"
    
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "CRM Upload 완료"

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet5.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet_조회_2.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_조회_2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "조회_2"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Long = 50000


Public Sub Sheet_조회_2_Clear()

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


Public Sub Sheet_조회_2_Query()

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

        Call Sheet_조회_2_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim i As Integer

        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.Cells(2, 6), "yyyyMMdd")
        
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        Select Case .Cells(3, 3)
                
            Case "최초 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_1").Offset(0, 0).Value2
        
            Case "최종 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_2").Offset(0, 0).Value2
        
            Case "개별 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_3").Offset(0, 0).Value2
        
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
        
        i = START_ROW_NUM
        
        While .Cells(i, 2) <> ""
        
            If .Cells(i, 17) > 0 Then
            
                Rows(i & ":" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.349986266670736
                    .PatternTintAndShade = 0
                End With
            
            End If
            
            i = i + 1
        
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
VBA MACRO Sheet_조회_1.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_조회_1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "조회_1"
Const Target_Sheet_Name As String = "CRM_입력"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Long = 50000


Public Sub Sheet_조회_1_Clear()

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


Public Sub Sheet_조회_1_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)

        If MsgBox("콜 기록을 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        '----------------------------------------------
        ' CRM_입력 시트로 복사가 안 된 데이터가 있는지 체크
        Dim i As Integer
        i = START_ROW_NUM
        
        While .Cells(i, 2) <> ""
        
            If .Cells(i, 2) = Date And .Cells(i, 1) <> "Y" Then
                MsgBox i & "번째 행이 아직 CRM_입력 시트 복사가 되지 않았습니다." & Chr(10) & Chr(10) & "CRM_입력 시트 복사를 한 후 조회 버튼을 눌러주세요."
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            i = i + 1
        
        Wend


        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Call Sheet_조회_1_Clear


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim ref_date_1 As String
        Dim ref_date_2 As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.Cells(2, 6), "yyyyMMdd")
        
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        
        Select Case .Cells(3, 3)
                
            Case "최초 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_1").Offset(0, 0).Value2
        
            Case "최종 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_2").Offset(0, 0).Value2
        
            Case "개별 콜 날짜"
                sql_str = Range("조회_1_과거_CALL_LOG_SELECT_SQL_3").Offset(0, 0).Value2
        
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
        
        i = START_ROW_NUM
        
        While .Cells(i, 2) <> ""
        
            If .Cells(i, 17) > 0 Then
            
                Rows(i & ":" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.349986266670736
                    .PatternTintAndShade = 0
                End With
            
            End If
            
            i = i + 1
        
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



Public Sub Sheet_조회_1_행추가()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
        
        Dim row_idx As Integer
        
        row_idx = ActiveCell.Row


        If row_idx < START_ROW_NUM Then
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If

        Sheets(This_Sheet_Name).Select
        Rows(row_idx & ":" & row_idx).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        .Cells(row_idx, 2) = Date
        .Cells(row_idx, 4) = Range("Cur_User_No").Offset(0, 1).Value2
        .Cells(row_idx, 5) = .Cells(row_idx + 1, 5)
        .Cells(row_idx, 6) = .Cells(row_idx + 1, 6)
        .Cells(row_idx, 7) = .Cells(row_idx + 1, 7)
        .Cells(row_idx, 8) = .Cells(row_idx + 1, 8)
        
        .Cells(row_idx, 15).Formula = "=IF(J" & row_idx & "="""","""",J" & row_idx & ")"
                
    
        Cells(row_idx, 9).Select

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    

End Sub




Public Sub Sheet_조회_1_CRM_입력_복사_개별()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
        
        Dim row_idx As Integer
        
        row_idx = ActiveCell.Row

        Sheets(This_Sheet_Name).Select
        
        
        If .Cells(row_idx, 2) <> Date Then
            MsgBox "당일 입력 call log가 아닙니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If .Cells(row_idx, 10) = "" Then
            MsgBox "콜 결과값이 없습니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If .Cells(row_idx, 1) = "Y" Then
            MsgBox "이미 CRM_입력 시트 복사가 된 행입니다. " & Chr(10) & Chr(10) & "그래도 복사하실려면 Y를 지우고 다시 해주세요."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If (.Cells(row_idx, 10) = "예약완료" Or .Cells(row_idx, 10) = "예약변경") And .Cells(row_idx, 11) = "" Then
            MsgBox "예약일이 비었습니다. 행번호 : " & row_idx
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
    
        If .Cells(row_idx, 11) <> "" And (.Cells(row_idx, 10) <> "예약완료" And .Cells(row_idx, 10) <> "예약변경") Then
            MsgBox "콜 결과값이 잘못 되었습니다. (예약일이 입력되어있음)  행번호 : " & row_idx
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If .Cells(row_idx, 4) = "" Then
            MsgBox "담당자가 없습니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If .Cells(row_idx, 5) = "" Then
            MsgBox "DB업체 칼럼 내용이 없습니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        
        If .Cells(row_idx, 7) = "" Then
            MsgBox "이름이 없습니다. 모르시는 경우 _ 를 입력해주세요."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        
        '--------------------------------------------------------------------------------------
        Dim target_row_idx As Integer
        
        target_row_idx = 7
        
        While Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) <> ""
            target_row_idx = target_row_idx + 1
        Wend
        '--------------------------------------------------------------------------------------
        
' 당일/직전 Calldate
'담당자
'DB업체
'DB정보
'이름
'연락처
'상담내용
'결과
'예약일
'예약시간
'내원여부
'담당자 메모
        
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) = .Cells(row_idx, 2)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 3) = .Cells(row_idx, 4)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 4) = .Cells(row_idx, 5)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 5) = IIf(.Cells(row_idx, 6) = "", "_", .Cells(row_idx, 6))
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 7)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = .Cells(row_idx, 8)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = .Cells(row_idx, 9)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 10)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = .Cells(row_idx, 11)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 12)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 13)
        Sheets(Target_Sheet_Name).Cells(target_row_idx, 13) = CStr(.Cells(row_idx, 14))
        
        'Cells(row_idx, 9).Select

        .Cells(row_idx, 1) = "Y"

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "복사 완료"

End Sub


Public Sub Sheet_조회_1_CRM_입력_복사_일괄()

End Sub




Public Sub Sheet_조회_1_내원여부_저장()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)
        
        Dim row_idx As Integer
        Dim user_no As String
        Dim sql_str As String
                
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""
        
        row_idx = ActiveCell.Row

        Sheets(This_Sheet_Name).Select
        
        
        If .Cells(row_idx, 10) <> "예약완료" And .Cells(row_idx, 10) <> "예약변경" Then
            MsgBox "콜 결과값이 예약완료 또는 예약변경이어야 합니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If .Cells(row_idx, 11) = "" Then
            MsgBox "예약일이 비었습니다."
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        End If
       
        
        '------------------------------------------------------------------------
        Dim select_db_agent As DB_Agent
        
        Set select_db_agent = New DB_Agent
        select_db_agent.Connect_DB
        
        
       '----------------------------------------------------------
        select_db_agent.Begin_Trans
        '----------------------------------------------------------
        
'Update
'    tm_call_log a
'set
'    a.VISITED_YN = :param06
',    a.CHART_NO = :param08
',    a.ALT_USER_NO = :param07
',    a.ALT_DATE   = to_char(sysdate, 'yyyymmdd')
',    a.ALT_TIME   = to_char(sysdate, 'hh24:mi:ss')
'where
'    1=1
'    and REF_DATE = :param01
'    and TM_NO = :param02
'    and DB_SRC_2 = :param03
'    and PHONE_NO = :param04
'    and SEQ_NO = :param05
      
        
        
        sql_str = Range("조회_1_내원여부_UPDATE_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, _
                            Format(.Cells(row_idx, 2), "yyyyMMdd"), _
                            user_no, _
                            .Cells(row_idx, 5), _
                            .Cells(row_idx, 8), _
                            .Cells(row_idx, 3), _
                            .Cells(row_idx, 13), _
                            user_no, _
                            .Cells(row_idx, 16))
        
        If Not select_db_agent.Insert_update_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
                
        '----------------------------------------------------------
        select_db_agent.Commit_Trans
        '----------------------------------------------------------
        
        Set select_db_agent = Nothing
        
        .Cells(row_idx, 1) = "U"

    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "저장 완료"

End Sub

-------------------------------------------------------------------------------
VBA MACRO Sheet7.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet7'
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
    
        If arglist(i) = "ALL" Then
            sql_str = WorksheetFunction.Substitute(sql_str, ":param" & Format(i + 1, "00"), "'%'")
        Else
            sql_str = WorksheetFunction.Substitute(sql_str, ":param" & Format(i + 1, "00"), "'" & arglist(i) & "'")
        End If
    
    Next i
    
    make_SQL = sql_str

End Function


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
    Range("B1212").Select
End Sub
Sub 매크로2()
Attribute 매크로2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로2 매크로
'

'
    Rows("1215:1215").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Rows("1215:1215").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Rows("1215:1215").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Rows("1216:1216").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Rows("1216:1216").Select

    Range("P1216").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P9").Select
    Selection.ClearContents

    Range("P8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P9").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P11").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P18").Select
    Selection.End(xlDown).Select
    Range("P1212").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("P1212").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B8").Select
End Sub
Sub 매크로3()
Attribute 매크로3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로3 매크로
'

'
 
    Range("P7").Select
End Sub
Sub 매크로4()
Attribute 매크로4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로4 매크로
'

'
    Range("P7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub
Sub 매크로5()
Attribute 매크로5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로5 매크로
'

'

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
    
    ' 다음으로 sheet별 intialize sub. call하기
    'Call Sheet_Deal_조회_Intialize_Button

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
        
                               
    Sheets("Config").Cells(11, 2) = emp_no
    Sheets("Config").Cells(11, 3) = emp_name

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
    
        If Sheets(p_sheet_name).Cells(i, p_col_idx) = p_name_to_find Then
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
    
    While Sheets(p_sheet_name).Cells(p_row_idx + i, p_col_idx) <> ""
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
         ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=R.Cells(PageRows * i + 1, 1)
      Next i
    Else
      nPagebreaks = 1
      For i = 1 To nPagebreaks
         ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=R.Cells(PageRows * i + 1, 1)
      Next i
    
    End If
    
    pages = ActiveSheet.HPageBreaks.Count
    pageBegin = "$A$1"
    For i = 1 To pages
      If i > 1 Then pageBegin = ActiveSheet.HPageBreaks(i - 1).Location.Address
      q = ActiveSheet.HPageBreaks(i).Location.Row - 1
      PrArea = pageBegin & ":" & "$H$" & Trim$(Str$(q))
      ActiveSheet.PageSetup.PrintArea = PrArea
      ActiveSheet.PageSetup.CenterFooter = Cells(q, 1)
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
VBA MACRO Sheet_당일배포.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet_당일배포'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Const This_Sheet_Name As String = "당일배포"
Const Target_Sheet_Name As String = "CRM_입력"

Public Const START_ROW_NUM As Integer = 8
Const MAX_ROW_NUM As Integer = 1000


Public Sub Sheet_당일배포_Clear()

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


Public Sub Sheet_당일배포_추가배포_조회()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    
    With Sheets(This_Sheet_Name)


        If MsgBox(.Cells(2, 3) & " 일자 배포 추가 배포 DB를 불러오시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            
            ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:= _
                True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                True
            Exit Sub
        
        End If


        If .Cells(5, 3) = 0 Then
        
            MsgBox "아직 배포조회를 하지 않았습니다. 추가배포 조회 전에 배포조회를 해주세요"
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            
            ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:= _
                True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                True
            Exit Sub
        
        End If


        
        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim last_ref_date As String
        Dim sql_str As String
        Dim user_no As String

        Dim cur_배포회차 As Integer
        
        cur_배포회차 = WorksheetFunction.Max(Range("C" & START_ROW_NUM & ":C" & MAX_ROW_NUM))

        

        row_idx = START_ROW_NUM
        REF_DATE = Format(.Cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("당일배포_추가배포_DB_SELECT_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE, user_no, cur_배포회차)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        While .Cells(row_idx, 2) <> ""
            row_idx = row_idx + 1
        Wend
        
        
        Range("B" & row_idx).CopyFromRecordset select_db_agent.DB_result_recordset

        .Cells(5, 3) = .Cells(5, 3) + select_db_agent.DB_result_recordset.RecordCount
        
        '----------------------------------------------------------------
        'row_idx = 8
        
        While .Cells(row_idx, 2) <> ""
        
            sql_str = Range("당일배포_기존_CALL_LOG_여부_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE, user_no, .Cells(row_idx, 7))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            If Not select_db_agent.DB_result_recordset.EOF Then
            
                If select_db_agent.DB_result_recordset(0) > 0 Then
                
                    Sheets(This_Sheet_Name).Select
                    Cells(row_idx, 7).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
                End If
            End If
        
            row_idx = row_idx + 1
        
        Wend
        
        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기

        row_idx = 8
        
        While .Cells(row_idx, 2) <> ""

            sql_str = Range("당일배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE, .Cells(row_idx, 7))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            If Not select_db_agent.DB_result_recordset.EOF Then
                   .Cells(row_idx, 18) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                     & "(" & select_db_agent.DB_result_recordset(1) & ")"

            End If
            
            
            If Left(.Cells(row_idx, 18), 2) = "취소" Then
                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                    select_db_agent.DB_result_recordset.MoveNext
                
                    Do While Not select_db_agent.DB_result_recordset.EOF
                        If select_db_agent.DB_result_recordset(8) = "미도래" Then
                            If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                .Cells(row_idx, 18) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
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

    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
       False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:= _
        True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
        True
    
    Range("B" & row_idx).Select
    
    MsgBox "조회 완료"

End Sub




Public Sub Sheet_당일배포_Query()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    
    With Sheets(This_Sheet_Name)


        If MsgBox(.Cells(2, 3) & " 일자 배포 DB를 불러오시겠습니까?" & Chr(10) & Chr(10) & "(아래 화면의 data는 화면에서 지워집니다)", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            
            ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:= _
                True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                True
            Exit Sub
        
        End If


        Call Sheet_당일배포_Clear

        .Cells(5, 3) = ""

        If Sheets(This_Sheet_Name).AutoFilterMode Then
            Rows("7:7").Select
            Selection.AutoFilter
        End If


        Dim select_db_agent As DB_Agent

        Dim row_idx As Integer
        Dim REF_DATE As String
        Dim last_ref_date As String
        Dim sql_str As String
        Dim user_no As String

        row_idx = START_ROW_NUM
        REF_DATE = Format(.Cells(2, 3), "yyyyMMdd")
        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("당일배포_배포DB_SELECT_SQL").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, REF_DATE, user_no)
        
        If Not select_db_agent.Select_DB(sql_str) Then
            MsgBox "ERROR"
        End If
        
        
        Range("B8").CopyFromRecordset select_db_agent.DB_result_recordset

        .Cells(5, 3) = select_db_agent.DB_result_recordset.RecordCount
        
        
        '----------------------------------------------------------------
        row_idx = 8
        
        While .Cells(row_idx, 2) <> ""
        
            sql_str = Range("당일배포_기존_CALL_LOG_여부_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE, user_no, .Cells(row_idx, 7))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            If Not select_db_agent.DB_result_recordset.EOF Then
            
                If select_db_agent.DB_result_recordset(0) > 0 Then
                
                    Sheets(This_Sheet_Name).Select
                    Cells(row_idx, 7).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    
                End If
            End If
        
            row_idx = row_idx + 1
        
        Wend
        
        ' ---------------------------------------------------------------------
        ' 덴트웹에서 예약상황 가져오기

        row_idx = 8
        
        While .Cells(row_idx, 2) <> ""

            sql_str = Range("당일배포_덴트웹_정보_SELECT_SQL").Offset(0, 0).Value2
            sql_str = make_SQL(sql_str, REF_DATE, .Cells(row_idx, 7))
        
            If Not select_db_agent.Select_DB(sql_str) Then
                MsgBox "ERROR"
            End If
            
            If Not select_db_agent.DB_result_recordset.EOF Then
                   .Cells(row_idx, 18) = IIf(select_db_agent.DB_result_recordset(9) > 0, "이행", select_db_agent.DB_result_recordset(8)) _
                                     & "(" & select_db_agent.DB_result_recordset(1) & ")"

            End If
            
            
            If Left(.Cells(row_idx, 18), 2) = "취소" Then
                If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                    select_db_agent.DB_result_recordset.MoveNext
                
                    Do While Not select_db_agent.DB_result_recordset.EOF
                        If select_db_agent.DB_result_recordset(8) = "미도래" Then
                            If Format(select_db_agent.DB_result_recordset(1), "yyyyMMdd") >= REF_DATE Then
                                .Cells(row_idx, 18) = select_db_agent.DB_result_recordset(8) & "(" & select_db_agent.DB_result_recordset(1) & ")"
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

    End With


    If Not Sheets(This_Sheet_Name).AutoFilterMode Then
        Rows("7:7").Select
        Selection.AutoFilter
    End If


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
       False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:= _
        True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
        True
    
    Range("B8").Select
    
    MsgBox "조회 완료"

End Sub



Public Sub Sheet_당일배포_CRM입력_복사()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With Sheets(This_Sheet_Name)


        If MsgBox("아래 콜 기록 전체를 CRM_입력 시트로 일괄 복사하시겠습니까?", vbYesNo + vbDefaultButton2) = vbNo Then
        
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Exit Sub
        
        End If


        Dim row_idx As Integer

        '--------------------------------------------------------------------------------------
        row_idx = START_ROW_NUM
        
        While .Cells(row_idx, 2) <> ""
        
            If .Cells(row_idx, 13) = "" Then
                
                MsgBox "Call 결과 값이 비었습니다. 행번호 : " & row_idx
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            If (.Cells(row_idx, 13) = "예약완료" Or .Cells(row_idx, 13) = "예약변경") And .Cells(row_idx, 14) = "" Then
            
                MsgBox "예약일이 비었습니다. 행번호 : " & row_idx
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If
        
            If .Cells(row_idx, 14) <> "" And (.Cells(row_idx, 13) <> "예약완료" And .Cells(row_idx, 13) <> "예약변경") Then
            
                MsgBox "콜 결과값이 잘못 되었습니다. (예약일이 입력되어있음)  행번호 : " & row_idx
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Exit Sub
            End If
        
        
            row_idx = row_idx + 1
        
        Wend
        '--------------------------------------------------------------------------------------


        Dim target_row_idx As Integer
        
        target_row_idx = 7
        
        While Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) <> ""
            target_row_idx = target_row_idx + 1
        Wend
        '--------------------------------------------------------------------------------------

        
        row_idx = START_ROW_NUM
        
        While .Cells(row_idx, 2) <> ""
        
' 당일/직전 Calldate
'담당자
'DB업체
'DB정보
'이름
'연락처
'상담내용
'결과
'예약일
'예약시간
'내원여부
'담당자 메모
        
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 2) = "당일"
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 3) = .Cells(row_idx, 2)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 4) = .Cells(row_idx, 4)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 5) = .Cells(row_idx, 5)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 6) = .Cells(row_idx, 6)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 7) = .Cells(row_idx, 7)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 8) = .Cells(row_idx, 12)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 9) = .Cells(row_idx, 13)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 10) = .Cells(row_idx, 14)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 11) = .Cells(row_idx, 15)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 12) = .Cells(row_idx, 16)
            Sheets(Target_Sheet_Name).Cells(target_row_idx, 13) = .Cells(row_idx, 17)
            
            row_idx = row_idx + 1
            target_row_idx = target_row_idx + 1
        
        Wend


    End With


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Range("B8").Select
    
    MsgBox "복사 완료"


End Sub



-------------------------------------------------------------------------------
VBA MACRO Sheet11.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet11'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet10.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet10'
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

'    Select Case Sheets(This_Sheet_Name).Cells(2, 8)
                
'        Case "TM별 일일 통계"
        
'            Call Sheet_통계_TM별_일일_통계_조회
        
'        Case "TM 기간 Performance(업체별)"
    
'            Call Sheet_통계_TM별_DB업체별_Performance_조회

'        Case Else

'            MsgBox "ERROR"
        
'    End Select

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
        
        Dim user_no As String

        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.Cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_일일TM통계_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2, user_no)
        
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

        While .Cells(row_idx, 2) <> ""
        
            If .Cells(row_idx, 4) = "합계" Then
            
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
        
        .Cells(7, 2) = "기준일"
        .Cells(7, 3) = "TM번호"
        .Cells(7, 4) = "TM이름"
        .Cells(7, 5) = "예약" & Chr(10) & "(당일DB)"
        .Cells(7, 6) = "예약" & Chr(10) & "(이전DB)"
        .Cells(7, 7) = "부재"
        .Cells(7, 8) = "타치과치료" & Chr(10) & "상담거부"
        .Cells(7, 9) = "상담완료" & Chr(10) & "재연락"
        .Cells(7, 10) = "결번"
        .Cells(7, 11) = "중복"
        .Cells(7, 12) = "본인아님" & Chr(10) & "신청안함"
        .Cells(7, 13) = "당일DB_개수"
        .Cells(7, 14) = "예약율"
        .Cells(7, 15) = "당일내원수"


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




Public Sub Sheet_통계_내원환자_리스트_조회()

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
        
        Dim user_no As String

        user_no = Range("Cur_User_No").Offset(0, 0).Value2
        
        
        Dim sql_str As String

        row_idx = START_ROW_NUM
        
        ref_date_1 = Format(.Cells(2, 3), "yyyyMMdd")
        ref_date_2 = Format(.Cells(2, 5), "yyyyMMdd")
        
        sql_str = ""

        Set select_db_agent = New DB_Agent
        
        sql_str = Range("통계_내원환자_리스트_SELECT").Offset(0, 0).Value2
        sql_str = make_SQL(sql_str, ref_date_1, ref_date_2, user_no)
        
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

        
        .Cells(7, 2) = "예약일"
        .Cells(7, 3) = "예약시간"
        .Cells(7, 4) = "차트번호"
        .Cells(7, 5) = "환자이름"
        .Cells(7, 6) = "전화번호"
        .Cells(7, 7) = "DB정보"
        .Cells(7, 8) = "DB업체"
        .Cells(7, 9) = "TM이름"

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




