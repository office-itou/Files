Rem ---------------------------------------------------------------------------
Rem ConvCovid19Data.vbs: CSV->TXT(TAB��؂�)�ϊ�
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
    Dim objFSO
    Dim CurDir

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    CurDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

Rem ---------------------------------------------------------------------------
    Dim InpDir
    Dim OutDir
    InpDir = CurDir & "\data"
    OutDir = CurDir & "\conv"
Rem ---------------------------------------------------------------------------
    Dim InpFileName
    Dim InpCount
    Dim InpLine
    Dim InpArray
    Dim InpDate
    Dim InpCode
    Dim InpName
    Dim InpValue
    Dim InpValue0
    Dim InpValue1

    Dim OutFileName
    Dim OutCount
    Dim OutLine
    Dim OutData()
    Dim OutValue()

    Dim OldCode
    Dim OldValue
Rem ---------------------------------------------------------------------------
    Dim Population()
    Dim RankData()
Rem ---------------------------------------------------------------------------
    Dim xlManual
    Dim xlAutomatic
    xlManual = -4135
    xlAutomatic = -4105

    Dim objExcel
    Dim objWorkbook
    Dim objWorksheet
    Dim objSrcWorkbook
    Dim objDstWorkbook
Rem ---------------------------------------------------------------------------
    Dim objFolder
    Dim objFile
Rem ---------------------------------------------------------------------------
    Dim I, J, K
    Dim Ret

Rem --- �l�� ------------------------------------------------------------------
    Wscript.Echo "�J�n�F�������f�[�^�["

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName="�l��(�l�����v2019).csv"
    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
Rem     InpLine = .ReadText(-2)
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
    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
Rem     InpLine = .ReadText(-2)
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

Rem Erase Population

    Wscript.Echo "�I���F�������f�[�^�["

Rem --- severe_cases_daily ----------------------------------------------------
    InpFileName="severe_cases_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutData
    ReDim OutData(1, 1999)

    OutData(0, 0) = "���t"
    OutData(1, 0) = "�d�ǎҐ�"

    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            OutData(0, InpCount + 1) = InpArray(0)
            OutData(1, InpCount + 1) = InpArray(1)
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "���o�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    Erase OutData

    Wscript.Echo "�I���F" & InpFileName

Rem --- requiring_inpatient_care_etc_daily ------------------------------------
    InpFileName="requiring_inpatient_care_etc_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutData
    ReDim OutData(3, 1999)

    OutData(0, 0) = "���t"
    OutData(1, 0) = "���@�Ґ�"
    OutData(2, 0) = "�މ@�Ґ�"
    OutData(3, 0) = "�m�F�\��"

    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        OldValue = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            OutData(0, InpCount + 1) = InpArray(0)
            OutData(1, InpCount + 1) = InpArray(1)
            If InpCount = 0 Then
                OldValue = InpArray(2)
            Else
                OutData(2, InpCount + 1) = InpArray(2) - OldValue
                OldValue = InpArray(2)
            End If
            OutData(3, InpCount + 1) = InpArray(3)
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "���o�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    Erase OutData

    Wscript.Echo "�I���F" & InpFileName

Rem --- pcr_case_daily --------------------------------------------------------
    InpFileName="pcr_case_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutData
    ReDim OutData(9, 1999)

    OutData(0, 0) = "���t"
    OutData(1, 0) = "���������ǌ�����"
    OutData(2, 0) = "���u��"
    OutData(3, 0) = "�n���q���������E�ی���"
    OutData(4, 0) = "���Ԍ�����Ёi��ɍs�������j"
    OutData(5, 0) = "��w��"
    OutData(6, 0) = "��Ë@��"
    OutData(7, 0) = "���v"
    OutData(8, 0) = "���Ԍ�����Ёi��Ɏ�����j"
    OutData(9, 0) = "�v"

    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            For I = 0 To 9
                OutData(I, InpCount + 1) = InpArray(I)
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "���o�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I) & Chr(9) & OutData(4, I) & Chr(9) & OutData(5, I) & Chr(9) & OutData(6, I) & Chr(9) & OutData(7, I) & Chr(9) & OutData(8, I) & Chr(9) & OutData(9, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    Erase OutData

    Wscript.Echo "�I���F" & InpFileName

Rem --- nhk_news_covid19_domestic_daily_data ----------------------------------
    InpFileName="nhk_news_covid19_domestic_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutData
    ReDim OutData(4, 1999)

    OutData(0, 0) = "���t"
    OutData(1, 0) = "���������Ґ�"
    OutData(2, 0) = "���������җ݌v"
    OutData(3, 0) = "�������ҎҐ�"
    OutData(4, 0) = "�������Ґ��݌v"

    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            OutData(0, InpCount + 1) = InpArray(0)
            OutData(1, InpCount + 1) = InpArray(1)
            OutData(2, InpCount + 1) = InpArray(2)
            OutData(3, InpCount + 1) = InpArray(3)
            OutData(4, InpCount + 1) = InpArray(4)
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "���o�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I) & Chr(9) & OutData(4, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

Rem Erase OutData

    Wscript.Echo "�I���F" & InpFileName

Rem --- nhk_news_covid19_prefectures_daily_data -------------------------------
    InpFileName="nhk_news_covid19_prefectures_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutValue
    ReDim OutValue(49, 1999, 4)
    Rem 0: �e�n�̊����Ґ�_1�����Ƃ̔��\��
    Rem 1: �e�n�̎��Ґ�_1�����Ƃ̔��\��
    Rem 2: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
    Rem 3: �e�n�̒���1�T�Ԃ̊����Ґ�(�Z�o)
    Rem 4: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�(�Z�o)

    For I = 0 To 4
        OutValue(0, 0, I) = "���t"
        OutValue(1, 0, I) = "�������v"
        OutValue(2, 0, I) = "��`���u�Ȃ�"
    Next

    Wscript.Echo "�Ǐo�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpLine = .ReadText(-2)
        InpCount = 0
        OldCode = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            InpDate = InpArray(0)       Rem ���t
            InpCode = InpArray(1)       Rem �s���{���R�[�h
            InpName = InpArray(2)       Rem �s���{����
Rem         InpArray(3)                 Rem �e�n�̊����Ґ�_1�����Ƃ̔��\��
Rem         InpArray(4)                 Rem �e�n�̊����Ґ�_�݌v
Rem         InpArray(5)                 Rem �e�n�̎��Ґ�_1�����Ƃ̔��\��
Rem         InpArray(6)                 Rem �e�n�̎��Ґ�_�݌v
Rem         InpArray(7)                 Rem �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
            If IsNumeric(InpCode) = True Then
                If OldCode <> InpCode Then
                    InpCount = 0
                    OldCode = InpCode
                    OutValue(InpCode + 2, 0, 0) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 1) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 2) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 3) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 4) = InpCode & ":" & InpName
                End If
                If IsDate(OutValue(0, InpCount + 1, 0)) = False Then
                    OutValue(0, InpCount + 1, 0) = InpDate
                    OutValue(0, InpCount + 1, 1) = InpDate
                    OutValue(0, InpCount + 1, 2) = InpDate
                    OutValue(0, InpCount + 1, 3) = InpDate
                    OutValue(0, InpCount + 1, 4) = InpDate
                End If
                OutValue(InpCode + 2, InpCount + 1, 0) = InpArray(3)
                OutValue(InpCode + 2, InpCount + 1, 1) = InpArray(5)
                OutValue(InpCode + 2, InpCount + 1, 2) = InpArray(7)
            End If
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "�v�Z�F" & InpFileName
    Rem --- "��`���u�Ȃ�"�̊����Ґ��E���Ґ�/�e�n�̒���1�T�Ԃ̊����Ґ��̎Z�o --
    For I = 0 To InpCount
        If CDate(OutData(0, I + 1)) = CDate(OutValue(0, I + 1, 0)) Then
            InpValue0 = 0
            InpValue1 = 0
            For J = 0 To 47 - 1
                InpValue0 = InpValue0 + OutValue(J + 3, I + 1, 0)
                InpValue1 = InpValue1 + OutValue(J + 3, I + 1, 1)
            Next
            OutValue(1, I + 1, 0) = OutData(1, I + 1)
            OutValue(1, I + 1, 1) = OutData(3, I + 1)
            OutValue(2, I + 1, 0) = OutData(1, I + 1) - InpValue0
            OutValue(2, I + 1, 1) = OutData(3, I + 1) - InpValue1
            If I >= 6 Then
                For J = 1 To 49
                    InpValue = 0
                    For K = 0 To 6
                        InpValue = InpValue + OutValue(J, I + 1 - K, 0)
                    Next
                    OutValue(J, I + 1, 3) = Round(InpValue / 7, 2)
                    Select Case J
                        Case 1          Rem �������v
                            If CDate(OutValue(0, I + 1, 0)) < CDate("2022/1/1") Then
                                OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J, 0) * 100000)
                            Else
                                OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J, 1) * 100000)
                            End If
                        Case 2          Rem ��`���u�Ȃ�
                        Case Else       Rem �e�n
                            If CDate(OutValue(0, I + 1, 0)) < CDate("2022/1/1") Then
                                OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J - 1, 0) * 100000)
                            Else
                                OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J - 1, 1) * 100000)
                            End If
                    End Select
                Next
            End If
        End If
    Next
    Rem --- �ŐV�̊e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ��̏��ʕt�� -------
    With CreateObject("ADODB.Recordset")
        .Fields.Append "CD",200,128
        .Fields.Append "NAME",200,128
        .Fields.Append "VALUE",5
        .Open
        For I = 0 To 46
            .AddNew
            .Fields("CD").Value = Left(OutValue(I + 3, 0, 4), 2)                Rem �s���{���R�[�h
            .Fields("NAME").Value = Mid(OutValue(I + 3, 0, 4), 4)               Rem �s���{����
            .Fields("VALUE").Value = OutValue(I + 3, InpCount + 0, 4)           Rem �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
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
            RankData(3, I) = FormatNumber(Round(.Fields("VALUE").Value, 2), 2, -1)
            RankData(4, I) = RankData(2, I) & ":" & RankData(3, I)
            .MoveNext
        Next
        .Close
    End With

    Wscript.Echo "���o�F" & InpFileName
    Rem --- 0: �e�n�̊����Ґ�_1�����Ƃ̔��\�� ---------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 49
                If OutLine = "" Then
                    OutLine = OutValue(J, I, 0)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I, 0)
                End If
            Next
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- 1: �e�n�̎��Ґ�_1�����Ƃ̔��\�� -----------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 49
                If OutLine = "" Then
                    OutLine = OutValue(J, I, 1)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I, 1)
                End If
            Next
            .WriteText OutLine, 1
        Next
        OutFileName = InpFileName & ".1.txt"
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- 2: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ� --------------------
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 49
                If OutLine = "" Then
                    OutLine = OutValue(J, I, 2)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I, 2)
                End If
            Next
            .WriteText OutLine, 1
        Next
        OutFileName = InpFileName & ".2.txt"
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- 3: �e�n�̒���1�T�Ԃ̊����Ґ� --------------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 49
                If OutLine = "" Then
                    OutLine = OutValue(J, I, 3)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I, 3)
                End If
            Next
            .WriteText OutLine, 1
        Next
        OutFileName = InpFileName & ".3.txt"
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- 4: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�(�Z�o) --------------
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 49
                If OutLine = "" Then
                    OutLine = OutValue(J, I, 4)
                Else
                    OutLine = OutLine & Chr(9) & OutValue(J, I, 4)
                End If
            Next
            .WriteText OutLine, 1
        Next
        OutFileName = InpFileName & ".4.txt"
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- �ŐV�̊e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ��̏��ʕt�� -------
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
        OutFileName = InpFileName & ".4.���ʕt��.txt"
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    Erase OutValue
    Erase OutData

    Wscript.Echo "�I���F" & InpFileName

Rem ---------------------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

Rem ---------------------------------------------------------------------------
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
    Set objFolder = objFSO.GetFolder(OutDir)
    For Each objFile in objFolder.files
        Wscript.Echo "���o�F" & objFile.name
        objExcel.Workbooks.OpenText objFile.Path,65001,,,,,True
        Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
        Select Case objFile.name
            Case "nhk_news_covid19_domestic_daily_data.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "����������"
            Case "nhk_news_covid19_prefectures_daily_data.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n������"
            Case "nhk_news_covid19_prefectures_daily_data.csv.1.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n���Ґ�"
            Case "nhk_news_covid19_prefectures_daily_data.csv.2.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n10���l"
            Case "nhk_news_covid19_prefectures_daily_data.csv.3.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n 7����"
            Case "nhk_news_covid19_prefectures_daily_data.csv.4.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n10���l(�Z�o)"
            Case "nhk_news_covid19_prefectures_daily_data.csv.4.���ʕt��.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n10���l(����)"
            Case "pcr_case_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "PCR ������"
            Case "requiring_inpatient_care_etc_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�������މ@"
            Case "severe_cases_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�����d�ǎ�"
        End Select
        With objExcel
            objSrcWorkbook.WorkSheets(1).Select
            .ActiveWindow.FreezePanes = False
            .Range("B2").Select
            .ActiveWindow.FreezePanes = True
        End With
        objSrcWorkbook.WorkSheets(1).Move ,objDstWorkbook.WorkSheets(objDstWorkbook.Sheets.Count)
        Set objSrcWorkbook = Nothing
    Next
    Wscript.Echo "�ۑ��F" & CurDir & "\Covid19Data.xlsx"
    objDstWorkbook.SaveAs(CurDir & "\Covid19Data.xlsx")

Rem ---------------------------------------------------------------------------
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
