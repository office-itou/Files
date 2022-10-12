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

    Dim obOrgjExcel
    Dim objOrgWorkbook
    Dim RowsEnd
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
Rem �d�ǎҐ��̐���
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

    Wscript.Echo "���o�F" & OutFileName
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
Rem ���@���Ó���v����ғ�����
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

    Wscript.Echo "���o�F" & OutFileName
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
Rem PCR�������{�l��
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

    Wscript.Echo "���o�F" & OutFileName
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

Rem --- �f�[�^�[�擾 ----------------------------------------------------------
    Erase OutValue
    ReDim OutValue(49, 1999, 4)
    Rem 0: �e�n�̊����Ґ�_1�����Ƃ̔��\��
    Rem 1: �e�n�̎��Ґ�_1�����Ƃ̔��\��
    Rem 2: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
    Rem 3: �e�n�̒���1�T�Ԃ̊����Ґ�(�Z�o)
    Rem 4: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�(�Z�o)
    Rem �^�C�g���s
    For I = 0 To 4
        OutValue(0, 0, I) = "���t"
        OutValue(1, 0, I) = "�������v"
        OutValue(2, 0, I) = "��`���u�Ȃ�"
        OutValue(2 +  1, 0, I) = "01:�k�C��"
        OutValue(2 +  2, 0, I) = "02:�X��"
        OutValue(2 +  3, 0, I) = "03:��茧"
        OutValue(2 +  4, 0, I) = "04:�{�錧"
        OutValue(2 +  5, 0, I) = "05:�H�c��"
        OutValue(2 +  6, 0, I) = "06:�R�`��"
        OutValue(2 +  7, 0, I) = "07:������"
        OutValue(2 +  8, 0, I) = "08:��錧"
        OutValue(2 +  9, 0, I) = "09:�Ȗ،�"
        OutValue(2 + 10, 0, I) = "10:�Q�n��"
        OutValue(2 + 11, 0, I) = "11:��ʌ�"
        OutValue(2 + 12, 0, I) = "12:��t��"
        OutValue(2 + 13, 0, I) = "13:�����s"
        OutValue(2 + 14, 0, I) = "14:�_�ސ쌧"
        OutValue(2 + 15, 0, I) = "15:�V����"
        OutValue(2 + 16, 0, I) = "16:�x�R��"
        OutValue(2 + 17, 0, I) = "17:�ΐ쌧"
        OutValue(2 + 18, 0, I) = "18:���䌧"
        OutValue(2 + 19, 0, I) = "19:�R����"
        OutValue(2 + 20, 0, I) = "20:���쌧"
        OutValue(2 + 21, 0, I) = "21:�򕌌�"
        OutValue(2 + 22, 0, I) = "22:�É���"
        OutValue(2 + 23, 0, I) = "23:���m��"
        OutValue(2 + 24, 0, I) = "24:�O�d��"
        OutValue(2 + 25, 0, I) = "25:���ꌧ"
        OutValue(2 + 26, 0, I) = "26:���s�{"
        OutValue(2 + 27, 0, I) = "27:���{"
        OutValue(2 + 28, 0, I) = "28:���Ɍ�"
        OutValue(2 + 29, 0, I) = "29:�ޗǌ�"
        OutValue(2 + 30, 0, I) = "30:�a�̎R��"
        OutValue(2 + 31, 0, I) = "31:���挧"
        OutValue(2 + 32, 0, I) = "32:������"
        OutValue(2 + 33, 0, I) = "33:���R��"
        OutValue(2 + 34, 0, I) = "34:�L����"
        OutValue(2 + 35, 0, I) = "35:�R����"
        OutValue(2 + 36, 0, I) = "36:������"
        OutValue(2 + 37, 0, I) = "37:���쌧"
        OutValue(2 + 38, 0, I) = "38:���Q��"
        OutValue(2 + 39, 0, I) = "39:���m��"
        OutValue(2 + 40, 0, I) = "40:������"
        OutValue(2 + 41, 0, I) = "41:���ꌧ"
        OutValue(2 + 42, 0, I) = "42:���茧"
        OutValue(2 + 43, 0, I) = "43:�F�{��"
        OutValue(2 + 44, 0, I) = "44:�啪��"
        OutValue(2 + 45, 0, I) = "45:�{�茧"
        OutValue(2 + 46, 0, I) = "46:��������"
        OutValue(2 + 47, 0, I) = "47:���ꌧ"
    Next

Rem --- deaths_cumulative_daily -----------------------------------------------
Rem ���S�Ґ��i�ݐρj
    InpFileName="deaths_cumulative_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutData
    ReDim OutData(49)

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
            For I = 0 To 48
                Select Case I
                    Case 0              Rem ���t
                        OutValue(0, InpCount + 1, 1) = InpArray(I)
                    Case 1              Rem �������v
                        If InpCount > 0 Then
                            OutValue(1, InpCount + 1, 1) = InpArray(I) - OutData(I)
                        Else
                            OutValue(1, InpCount + 1, 1) = InpArray(I)
                        End If
                        OutData(I) = InpArray(I)
                    Case Else           Rem �e�n��1�����Ƃ̔��\��
                        If InpCount > 0 Then
                            OutValue(1 + I, InpCount + 1, 1) = InpArray(I) - OutData(1 + I)
                        Else
                            OutValue(1 + I, InpCount + 1, 1) = InpArray(I)
                        End If
                        OutData(1 + I) = InpArray(I)
                End Select
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Rem --- �e�n�̎��Ґ�_1�����Ƃ̔��\�� --------------------------------------
    Wscript.Echo "���o�F" & OutFileName

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
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

Rem --- newly_confirmed_cases_daily -------------------------------------------
Rem �V�K�z���Ґ��̐��ځi���ʁj
    InpFileName="newly_confirmed_cases_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

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
            For I = 0 To 48
                Select Case I
                    Case 0          Rem ���t
                        OutValue(0, InpCount + 1, 0) = InpArray(I)
                        OutValue(0, InpCount + 1, 3) = InpArray(I)
                        OutValue(0, InpCount + 1, 4) = InpArray(I)
                    Case 1          Rem �������v
                        OutValue(1, InpCount + 1, 0) = InpArray(I)
                    Case Else       Rem �e�n��1�����Ƃ̔��\��
                        OutValue(1 + I, InpCount + 1, 0) = InpArray(I)
                End Select
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Rem --- �e�n�̒���1�T�Ԃ̊����Ґ��̎Z�o -----------------------------------
    Wscript.Echo "�v�Z�F" & InpFileName
    For I = 0 To InpCount
        If I >= 6 Then
            For J = 1 To 49
                InpValue = 0
                Rem 7���ԍ��v
                For K = 0 To 6
                    InpValue = InpValue + OutValue(J, I + 1 - K, 0)
                Next
                Rem 7���ԕ���
                OutValue(J, I + 1, 3) = Round(InpValue / 7, 2)
                Rem �l��10���l������
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

    Rem --- 0: �e�n�̊����Ґ�_1�����Ƃ̔��\�� ---------------------------------
    Wscript.Echo "���o�F" & OutFileName
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
    Rem --- 3: �e�n�̒���1�T�Ԃ̊����Ґ� --------------------------------------
    OutFileName = InpFileName & ".3.txt"
    Wscript.Echo "���o�F" & OutFileName
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
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- 4: �e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�(�Z�o) --------------
    OutFileName = InpFileName & ".4.txt"
    Wscript.Echo "���o�F" & OutFileName
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
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With
    Rem --- �ŐV�̊e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ��̏��ʕt�� -------
    OutFileName = InpFileName & ".4.���ʕt��.txt"
    Wscript.Echo "���o�F" & OutFileName
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
            Case "newly_confirmed_cases_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n������"
            Case "deaths_cumulative_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n���Ґ�"
            Case "newly_confirmed_cases_daily.csv.3.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n 7����"
            Case "newly_confirmed_cases_daily.csv.4.txt"
                objSrcWorkbook.WorkSheets(1).Name = "�e�n10���l(�Z�o)"
            Case "newly_confirmed_cases_daily.csv.4.���ʕt��.txt"
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
Rem �O���t      �F�e�n10���l(����)
Rem �����Ґ�    �F�e�n������
Rem 7���ԕ��ϒl �F�e�n 7����
Rem 10���l������F�e�n10���l/�e�n10���l(�Z�o)
Rem ���{����    �F����������/�����d�ǎ�/�������މ@/PCR ������
    Set obOrgjExcel = GetObject(, "Excel.Application")
    For Each objOrgWorkbook In obOrgjExcel.Workbooks
        If objOrgWorkbook.Name = "covid-19.xlsx" Then
            Ret = MsgBox("�f�[�^�[�̃R�s�[�����܂����H", vbYesNo)
            If Ret = 6 Then
                Wscript.Echo "�]���FCovid19Data.xlsx��covid-19.xlsx"
                Rem --- �����Ґ� ----------------------------------------------
                With objDstWorkbook.WorkSheets("�e�n������")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("�����Ґ�").Range("B" & (RowsEnd + 3 - 2) & ":AX" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":AX" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("�����Ґ�")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- 7���ԕ��ϒl -------------------------------------------
                With objDstWorkbook.WorkSheets("�e�n 7����")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
Rem                 objOrgWorkbook.WorkSheets("7���ԕ��ϒl").Range("B" & (RowsEnd + 3 - 2) & ":AX" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":AX" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("7���ԕ��ϒl")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- 10���l������ ------------------------------------------
                With objDstWorkbook.WorkSheets("�e�n10���l(�Z�o)")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
Rem                 objOrgWorkbook.WorkSheets("10���l������").Range("B" & (RowsEnd + 3 - 2) & ":B" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("10���l������")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- ���{���� ----------------------------------------------
                With objDstWorkbook.WorkSheets("�����d�ǎ�")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("G" & (RowsEnd + 989 - 874) & ":G" & (RowsEnd + 989 - 874)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("�������މ@")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("I" & (RowsEnd + 117 - 2) & ":I" & (RowsEnd + 117 - 2)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                    RowsEnd = .Cells(.Rows.Count, 3).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("J" & (RowsEnd + 118 - 3) & ":J" & (RowsEnd + 118 - 3)).Value = .Range("C" & RowsEnd & ":C" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("PCR ������")
                    RowsEnd = .Cells(.Rows.Count, 10).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("K" & (RowsEnd + 36 - 2) & ":K" & (RowsEnd + 36 - 2)).Value = .Range("J" & RowsEnd & ":J" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("�e�n���Ґ�")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("���{����").Range("E" & (RowsEnd + 989 - 874) & ":E" & (RowsEnd + 989 - 874)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("���{����")
                    .Activate 
                    .Range("B" & (RowsEnd + 989 - 874)).Select
                End With
                Rem --- �O���t�p ----------------------------------------------
Rem             With objDstWorkbook.WorkSheets("�e�n10���l(����)")
Rem                 RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
Rem                 .Range("E2:E" & RowsEnd).Copy
Rem             End With
                Rem --- �I������ ----------------------------------------------
                objDstWorkbook.Close
                objExcel.Quit
            End If
            Exit For
        End If
    Next
    Set obOrgjExcel = Nothing

Rem ---------------------------------------------------------------------------
    Set objDstWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
