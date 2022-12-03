Rem ---------------------------------------------------------------------------
Rem 
Rem ---------------------------------------------------------------------------
Option Explicit

Rem ---------------------------------------------------------------------------
    Dim objShell
    Dim Arguments

    If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
        For I = 0 To WScript.Arguments.Count - 1
            Arguments = Arguments & " """ & WScript.Arguments.Item(I) & """"
        Next
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "CScript """ & WScript.ScriptFullName & """ " & Arguments
        Set objShell =  Nothing
        WScript.Quit
    End If

Rem ---------------------------------------------------------------------------
    Dim Ret
    Dim I, J, K
    Rem -----------------------------------------------------------------------
    Dim objFSO
    Dim objFolder
    Dim objFile
    Dim CurDir
    Dim InpDir
    Dim OutDir
    Dim InpFileName
    Dim OutFileName
    Rem -----------------------------------------------------------------------
    Dim InpCount
    Dim InpLine
    Dim InpArray
    Dim InpValue
    Dim OutCount
    Dim OutLine
    Dim OutValue()
    Dim OutData
    Rem -----------------------------------------------------------------------
    Dim objExcel
    Dim objWorkbook
    Dim objWorksheet
    Dim objSrcWorkbook
    Dim objDstWorkbook
    Dim WorkSheetName
    Dim Target
    Rem -----------------------------------------------------------------------
    Dim obOrgjExcel
    Dim objOrgWorkbook
    Dim RowsEnd
    Rem -----------------------------------------------------------------------
    Dim Population()
    Rem -----------------------------------------------------------------------
    Dim DateList()
    Dim RankData()

Rem ---------------------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    CurDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
    InpDir = CurDir & "\data"
    OutDir = CurDir & "\conv"

Rem --- �l�� ------------------------------------------------------------------
    WScript.Echo "�J�n�F�������f�[�^�["

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName="�l��(�l�����v2019).csv"
    WScript.Echo "�Ǐo�F" & InpFileName

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            For I = 0 To 3
                Population(I, InpCount, 0) = InpArray(I)
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    InpFileName="�l��(��������2020).csv"
    WScript.Echo "�Ǐo�F" & InpFileName

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            For I = 0 To 3
                Population(I, InpCount, 1) = InpArray(I)
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    WScript.Echo "�I���F�������f�[�^�["

Rem --- �f�[�^�[�ϊ� ----------------------------------------------------------
    WScript.Echo "�J�n�F�f�[�^�[�ϊ�"

    Erase DateList
    ReDim DateList(1999)

    DateList(0) = "���t"

    Rem --- �s���{�����Ƃ̊����Ґ�[����/7������/10���l������] -----------------
    Erase OutValue
    ReDim OutValue(48, 1999, 4)

    For I = 0 To 4
        OutValue(0     , 0, I) = "���t"
        OutValue(1 +  0, 0, I) = "�������v"
        OutValue(1 +  1, 0, I) = "�k�C��"
        OutValue(1 +  2, 0, I) = "�X��"
        OutValue(1 +  3, 0, I) = "��茧"
        OutValue(1 +  4, 0, I) = "�{�錧"
        OutValue(1 +  5, 0, I) = "�H�c��"
        OutValue(1 +  6, 0, I) = "�R�`��"
        OutValue(1 +  7, 0, I) = "������"
        OutValue(1 +  8, 0, I) = "��錧"
        OutValue(1 +  9, 0, I) = "�Ȗ،�"
        OutValue(1 + 10, 0, I) = "�Q�n��"
        OutValue(1 + 11, 0, I) = "��ʌ�"
        OutValue(1 + 12, 0, I) = "��t��"
        OutValue(1 + 13, 0, I) = "�����s"
        OutValue(1 + 14, 0, I) = "�_�ސ쌧"
        OutValue(1 + 15, 0, I) = "�V����"
        OutValue(1 + 16, 0, I) = "�x�R��"
        OutValue(1 + 17, 0, I) = "�ΐ쌧"
        OutValue(1 + 18, 0, I) = "���䌧"
        OutValue(1 + 19, 0, I) = "�R����"
        OutValue(1 + 20, 0, I) = "���쌧"
        OutValue(1 + 21, 0, I) = "�򕌌�"
        OutValue(1 + 22, 0, I) = "�É���"
        OutValue(1 + 23, 0, I) = "���m��"
        OutValue(1 + 24, 0, I) = "�O�d��"
        OutValue(1 + 25, 0, I) = "���ꌧ"
        OutValue(1 + 26, 0, I) = "���s�{"
        OutValue(1 + 27, 0, I) = "���{"
        OutValue(1 + 28, 0, I) = "���Ɍ�"
        OutValue(1 + 29, 0, I) = "�ޗǌ�"
        OutValue(1 + 30, 0, I) = "�a�̎R��"
        OutValue(1 + 31, 0, I) = "���挧"
        OutValue(1 + 32, 0, I) = "������"
        OutValue(1 + 33, 0, I) = "���R��"
        OutValue(1 + 34, 0, I) = "�L����"
        OutValue(1 + 35, 0, I) = "�R����"
        OutValue(1 + 36, 0, I) = "������"
        OutValue(1 + 37, 0, I) = "���쌧"
        OutValue(1 + 38, 0, I) = "���Q��"
        OutValue(1 + 39, 0, I) = "���m��"
        OutValue(1 + 40, 0, I) = "������"
        OutValue(1 + 41, 0, I) = "���ꌧ"
        OutValue(1 + 42, 0, I) = "���茧"
        OutValue(1 + 43, 0, I) = "�F�{��"
        OutValue(1 + 44, 0, I) = "�啪��"
        OutValue(1 + 45, 0, I) = "�{�茧"
        OutValue(1 + 46, 0, I) = "��������"
        OutValue(1 + 47, 0, I) = "���ꌧ"
    Next
    Rem -----------------------------------------------------------------------
    InpFileName = "newly_confirmed_cases_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            For I = 0 To 48
                If I = 0 Then
                    DateList(   InpCount + 1   ) = InpArray(I)                  Rem ���t�ꗗ
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  Rem ����
                    OutValue(I, InpCount + 1, 1) = InpArray(I)                  Rem 7������
                    OutValue(I, InpCount + 1, 2) = InpArray(I)                  Rem 10���l������
                    OutValue(I, InpCount + 1, 3) = InpArray(I)                  Rem 10���l������(7���ԕ���)
                    OutValue(I, InpCount + 1, 4) = InpArray(I)                  Rem 10���l������(�Z�o)
                Else
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  Rem ����
                    If InpCount >= 6 Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 0)
                        Next
                        OutValue(I, InpCount + 1, 1) = Round(InpValue / 7, 2)   Rem 7���ԕ���
                                                                                Rem 10���l������(�Z�o)
                        If CDate(OutValue(0, InpCount + 1, 0)) < CDate("2022/1/1") Then
                            OutValue(I, InpCount + 1, 4) = CDbl(InpValue / Population(3, I, 0) * 100000)
                        Else
                            OutValue(I, InpCount + 1, 4) = CDbl(InpValue / Population(3, I, 1) * 100000)
                        End If
                    End If
                End If
            Next
            InpCount = InpCount + 1
        Loop
        OutCount = InpCount
        .Close
    End With
    Rem -----------------------------------------------------------------------
    InpFileName = "newly_confirmed_cases_per_100_thousand_population_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I, 2)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            For I = 1 To 48
                OutValue(I, InpCount + 1, 2) = InpArray(I)                      Rem 10���l������(����)
                If InpCount >= 6 Then
                    If IsNumeric(OutValue(I, InpCount + 1 - 7, 2)) = True Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 2)
                        Next
                        OutValue(I, InpCount + 1, 3) = InpValue                 Rem 10���l������(7���ԍ��v)
                    End If
                End If
            Next
            InpCount = InpCount + 1
        Loop
        OutCount = InpCount
        .Close
    End With
    Rem -----------------------------------------------------------------------
    For I = 0 To 4
        OutFileName = "�����Ґ�." & I & ".txt"
        WScript.Echo "���o�F" & OutFileName
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            For J = 0 To OutCount
                OutLine = ""
                For K = 0 To 48
                    If OutLine = "" Then
                        OutLine = OutValue(K, J, I)
                    Else
                        OutLine = OutLine & Chr(9) & OutValue(K, J, I)
                    End If
                Next
                .WriteText OutLine, 1
            Next
            .SaveToFile OutDir & "\" & OutFileName, 2
            .Close
        End With
    Next
    Rem --- ���ʕt�� ----------------------------------------------------------
    With CreateObject("ADODB.Recordset")
        .Fields.Append "CD",200,128
        .Fields.Append "NAME",200,128
        .Fields.Append "VALUE",5
        .Open
        For I = 0 To 46
            .AddNew
            .Fields("CD").Value = I + 1                                         Rem �s���{���R�[�h
            .Fields("NAME").Value = OutValue(I + 2, 0, 4)                       Rem �s���{����
            .Fields("VALUE").Value = OutValue(I + 2, InpCount + 0, 4)           Rem �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
        Next
        .Sort = "VALUE DESC,CD"
        .MoveFirst
        Erase RankData
        ReDim RankData(4, 47)
        RankData(0, 0) = "����"
        RankData(1, 0) = "�s���{��CD"
        RankData(2, 0) = "�s���{����"
        RankData(3, 0) = "�����Ґ�"
        RankData(4, 0) = "�R�s�y�p"
        For I = 1 To 47
            RankData(0, I) = I
            RankData(1, I) = .Fields("CD").Value
            RankData(2, I) = .Fields("NAME").Value
            RankData(3, I) = FormatNumber(Round(.Fields("VALUE").Value, 2), 2, -1, 0, 0)
            RankData(4, I) = RankData(2, I) & ":" & RankData(3, I)
            .MoveNext
        Next
        .Close
    End With
    Rem -----------------------------------------------------------------------
    OutFileName = "���ʕt��.txt"
    WScript.Echo "���o�F" & OutFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To 47
            OutLine = RankData(0, I) & _
             Chr(9) & RankData(1, I) & _
             Chr(9) & RankData(2, I) & _
             Chr(9) & RankData(3, I) & _
             Chr(9) & RankData(4, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    Rem --- ���{����[�����Ґ�/���Ґ�/�d�ǎҐ�/���@�×{��/�މ@�×{����/PCR������]
    Erase OutValue
    ReDim OutValue(16, 1999)

    OutValue( 0, 0) = "���t"
    OutValue( 1, 0) = "�����Ґ�"
    OutValue( 2, 0) = "�����Ґ�(7���ԕ���)"
    OutValue( 3, 0) = "�����Ґ�(10���l������)"
    OutValue( 4, 0) = "�����Ґ�(10���l������E7���ԕ���)"
    OutValue( 5, 0) = "�����Ґ�(�݌v)"
    OutValue( 6, 0) = "���Ґ�"
    OutValue( 7, 0) = "���Ґ�(7���ԕ���)"
    OutValue( 8, 0) = "���Ґ�(�݌v)"
    OutValue( 9, 0) = "�d�ǎҐ�"
    OutValue(10, 0) = "�d�ǎҐ�(7���ԕ���)"
    OutValue(11, 0) = "���@�×{��"
    OutValue(12, 0) = "�މ@�×{����"
    OutValue(13, 0) = "�މ@�×{����(�݌v)"
    OutValue(14, 0) = "PCR������"
    OutValue(15, 0) = "�z����"
    OutValue(16, 0) = "�z����(7���ԕ��ϒl)"

    For I = 0 To OutCount
        OutValue( 0, I + 1) = DateList(I + 1)
    Next

    Rem --- �����Ґ�[�݌v] ----------------------------------------------------
    InpFileName = "confirmed_cases_cumulative_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
                                                           Rem �����Ґ�[�����E�S��]
            If IsNumeric(OutValue(4, InpCount + 1 - 1)) = False Then
                OutValue(1, InpCount + 1) = InpArray(1)
            Else
                OutValue(1, InpCount + 1) = InpArray(1) - OutValue(5, InpCount + 1 - 1)
            End If
                                                                                Rem 10���l������(����)
            If CDate(OutValue(0, InpCount + 1)) < CDate("2022/1/1") Then
                OutValue(3, InpCount + 1) = CDbl(OutValue(1, InpCount + 1) / Population(3, 1, 0) * 100000)
            Else
                OutValue(3, InpCount + 1) = CDbl(OutValue(1, InpCount + 1) / Population(3, 1, 1) * 100000)
            End If
            If InpCount >= 6 Then
                InpValue = 0
                For I = 0 To 6
                    InpValue = InpValue + OutValue(1, InpCount + 1 - I)
                Next
                OutValue(2, InpCount + 1) = Round(InpValue / 7, 2)              Rem 7���ԕ���
                If IsNumeric(OutValue(3, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For J = 0 To 6
                        InpValue = InpValue + OutValue(3, InpCount + 1 - J)
                    Next
                    OutValue(4, InpCount + 1) = InpValue                        Rem 10���l������(7���ԍ��v)
                End If
            End If

            OutValue(5, InpCount + 1) = InpArray(1)        Rem �����Ґ�[�݌v�E�S��]
            InpCount = InpCount + 1
        Loop
        OutCount = InpCount
        .Close
    End With
    Rem --- ���Ґ� ------------------------------------------------------------
    InpFileName = "deaths_cumulative_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        OutData = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
                                                            Rem ���Ґ�[���ʁE�S��]
            If IsNumeric(OutValue(6, InpCount + 1 - 1)) = False Then
                OutValue(6, InpCount + 1) = InpArray(1)
            Else
                OutValue(6, InpCount + 1) = InpArray(1) - OutData
            End If
            OutData = InpArray(1)
                                                            Rem 7���ԕ���
            If InpCount >= 6 Then
                If IsNumeric(OutValue(6, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For I = 0 To 6
                        InpValue = InpValue + OutValue(6, InpCount + 1 - I)
                    Next
                    OutValue(7, InpCount + 1) = Round(InpValue / 7, 2)
                End If
            End If
            OutValue(8, InpCount + 1) = InpArray(1)        Rem ���Ґ�[�݌v�E�S��]
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- �d�ǎҐ� ----------------------------------------------------------
    InpFileName = "severe_cases_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            OutValue(9, InpCount + 1) = InpArray(1)         Rem �d�ǎҐ�[���ʁE�S��]
                                                            Rem 7���ԕ���
            If InpCount >= 6 Then
                If IsNumeric(OutValue(9, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For I = 0 To 6
                        InpValue = InpValue + OutValue(9, InpCount + 1 - I)
                    Next
                    OutValue(10, InpCount + 1) = Round(InpValue / 7, 2)
                End If
            End If
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- ���@�×{��/�މ@�×{���� -------------------------------------------
    InpFileName = "requiring_inpatient_care_etc_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        OutData = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            OutValue(11, InpCount + 1) = InpArray(1)        Rem ���@�×{��[���ʁE�S��]
                                                            Rem �މ@�×{����[���ʁE�S��]
            If IsNumeric(OutValue(12, InpCount + 1 - 1)) = False Then
                OutValue(12, InpCount + 1) = InpArray(2)
            Else
                OutValue(12, InpCount + 1) = InpArray(2) - OutData
            End If
            OutValue(13, InpCount + 1) = InpArray(2)        Rem �މ@�×{����[�݌v�E�S��]
            OutData = InpArray(2)
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- PCR������ ---------------------------------------------------------
    InpFileName = "pcr_case_daily.csv"
    WScript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount = 0 Then
                For I = 1 To OutCount
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            OutValue(14, InpCount + 1) = InpArray(9)        Rem PCR������
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- ���{���� ----------------------------------------------------------
    OutFileName = "���{����.txt"
    WScript.Echo "���o�F" & OutFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To OutCount
            OutLine = ""
            For J = 0 To 14
                If OutLine = "" Then
                    OutLine = OutValue(J, I)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I)
                End If
            Next
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem -----------------------------------------------------------------------
    Erase OutValue

    WScript.Echo "�I���F�f�[�^�[�ϊ�"

Rem --- Excel -----------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

    Rem -----------------------------------------------------------------------
    MakeExcelFile "�����Ґ�", "�����Ґ�.0.txt"
    MakeExcelFile "7������" , "�����Ґ�.1.txt"
    MakeExcelFile "10���l"  , "�����Ґ�.3.txt"
    MakeExcelFile "���{����", "���{����.txt"
    MakeExcelFile "���ʕt��", "���ʕt��.txt"

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("Sheet1")
        .Activate
        .Name = "�O���t"
    End With

    Rem -----------------------------------------------------------------------
    WScript.Echo "�ۑ��F" & CurDir & "\" & "MakeCovid19Graph.xlsx"
    objDstWorkbook.SaveAs(CurDir & "\" & "MakeCovid19Graph.xlsx")

    Rem -----------------------------------------------------------------------
    Set obOrgjExcel = GetObject(, "Excel.Application")
    For Each objOrgWorkbook In obOrgjExcel.Workbooks
        If objOrgWorkbook.Name = "covid-19.xlsx" Then
            Ret = MsgBox("�f�[�^�[�̃R�s�[�����܂����H", vbYesNo)
            If Ret = 6 Then
                Wscript.Echo "�]���FCovid19Data.xlsx��covid-19.xlsx"
                Rem --- �����Ґ� ----------------------------------------------
                With objDstWorkbook.WorkSheets("�����Ґ�")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("�����Ґ�").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("�����Ґ�")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 7������ -----------------------------------------------
                With objDstWorkbook.WorkSheets("7������")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("7������").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("7������")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 10���l ------------------------------------------------
                With objDstWorkbook.WorkSheets("10���l")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("10���l").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("10���l")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- ���{���� ----------------------------------------------
                With objDstWorkbook.WorkSheets("���{����")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("A1:O" & RowsEnd).Value = .Range("A1:O" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("���{����")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- ���ʕt�� ----------------------------------------------
                With objDstWorkbook.WorkSheets("���ʕt��")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���ʕt��").Range("A1:E" & RowsEnd).Value = .Range("A1:E" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("���ʕt��")
                    .Activate 
                    .Range("A1").Select
                End With
                Rem --- �O���t�p ----------------------------------------------
                With objDstWorkbook.WorkSheets("���ʕt��")
                    RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
                    .Range("E2:E" & RowsEnd).Copy
                End With
                Rem --- �I������ ----------------------------------------------
                With objOrgWorkbook.WorkSheets("�O���t")
                    .Activate
                    .Range("A1").Select
                End With
                objDstWorkbook.Close
                objExcel.Quit
            End If
            Exit For
        End If
    Next

    Rem -----------------------------------------------------------------------
    Set objDstWorkbook = Nothing
    Set objExcel = Nothing

Rem ---------------------------------------------------------------------------
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
Rem WScript.Quit

Rem ---------------------------------------------------------------------------
Sub MakeExcelFile(WorkSheetName, InpFileName)
    WScript.Echo "���o�F" & InpFileName
    objExcel.Workbooks.OpenText OutDir & "\" & InpFileName,65001,,,,,True
    Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
    objSrcWorkbook.WorkSheets(1).Name = WorkSheetName
    With objExcel
        objSrcWorkbook.WorkSheets(1).Select
        .ActiveWindow.FreezePanes = False
        .Range("B2").Select
        .ActiveWindow.FreezePanes = True
        Select Case WorkSheetName
            Case "�����Ґ�"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
            Case "7������"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "10���l"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "���{����"
                With .Range("A1:N1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:B" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("C2:C" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
                .Range("D2:D" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
                .Range("E2:E" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
                .Range("F2:F" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("G2:G" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("H2:H" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
                .Range("I2:I" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("J2:J" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("K2:K" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
                .Range("L2:L" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("M2:M" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("N2:N" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("O2:O" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
            Case "���ʕt��"
                With .Range("A1:E1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("D2:D" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
        End Select
        If .Range("A1").Text = "���t" Then
            With .Range("A2:A" & (.Cells(.Rows.Count, 1).End(-4162).Row))
                .NumberFormatLocal = "yyyy/mm/dd(aaa)"
                .HorizontalAlignment = -4108
            End With
            .Columns("A").AutoFit
        End If
        .Cells.EntireRow.AutoFit
        If .Range("A1").Text = "���t" Then
            .Range("B" & (.Cells(.Rows.Count, 1).End(-4162).Row)).Select
        Else
            .Range("B2").Select
        End If
    End With
    objSrcWorkbook.WorkSheets(1).Move ,objDstWorkbook.WorkSheets(objDstWorkbook.Sheets.Count)
    Set objSrcWorkbook = Nothing
End Sub

Rem --- Memo ------------------------------------------------------------------
Rem objExcel.Workbooks.OpenText FileName, _
Rem                             Origin, _
Rem                             StartRow, _
Rem                             DataType, _
Rem                             TextQualifier, _
Rem                             ConsecutiveDelimiter, _
Rem                             Tab, _
Rem                             Semicolon, _
Rem                             Comma, _
Rem                             Space, _
Rem                             Other, _
Rem                             OtherChar, _
Rem                             FieldInfo, _
Rem                             TextVisualLayout, _
Rem                             DecimalSeparator, _
Rem                             ThousandsSeparator, _
Rem                             TrailingMinusNumbers, _
Rem                             Local

Rem --- EOF -------------------------------------------------------------------
