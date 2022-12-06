Rem ---------------------------------------------------------------------------
Rem
Rem ---------------------------------------------------------------------------
Option Explicit
On Error GoTo 0

Rem ---------------------------------------------------------------------------
    Dim objShell
    Dim Arguments

    If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
        For I = 0 To WScript.Arguments.Count - 1
            Arguments = Arguments & " """ & WScript.Arguments.Item(I) & """"
        Next
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "CScript """ & WScript.ScriptFullName & """ " & Arguments
        Set objShell = Nothing
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
    Dim PicDir
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
    Dim OutDate
    Rem -----------------------------------------------------------------------
    Dim objExcel
    Dim objWorkbook
    Dim objWorksheet
    Dim objSrcWorkbook
    Dim objDstWorkbook
    Dim WorkSheetName
    Dim Target
    Rem -----------------------------------------------------------------------
    Dim objChart
    Dim Charts
    Dim objRange
    Dim objRange2
    Rem -----------------------------------------------------------------------
    Dim objOrgExcel
    Dim objOrgWorkbook
    Dim RowsEnd
    Dim MaxRow
    Dim MinRow
    Dim MaxColumn
    Dim MinColumn
    Dim MaxValue
    Dim MinValue
    Dim PosiCD
    Dim PosiTop
    Dim PosiLeft
    Dim PointTop
    Rem -----------------------------------------------------------------------
    Dim Population()
    Rem -----------------------------------------------------------------------
    Dim DateList()
    Dim RankData()
    Rem -----------------------------------------------------------------------
    Dim Collection()

Rem ---------------------------------------------------------------------------
    WScript.Echo Now

Rem ---------------------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    CurDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
    InpDir = CurDir & "\data"
    OutDir = CurDir & "\conv"
    PicDir = CurDir & "\grph"

Rem --- �l�� ------------------------------------------------------------------
    WScript.Echo "�J�n�F�������f�[�^�["

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName = "�l��(�l�����v2019).csv"
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

    InpFileName = "�l��(��������2020).csv"
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

    Rem --- �s���{�����Ƃ̊����Ґ�[����/7���ԕ���/10���l������] -----------------
    Erase OutValue
    ReDim OutValue(48, UBound(DateList), 4)

    For I = 0 To 4
        OutValue(0, 0, I) = "���t"
        OutValue(1 + 0, 0, I) = "�������v"
        OutValue(1 + 1, 0, I) = "�k�C��"
        OutValue(1 + 2, 0, I) = "�X��"
        OutValue(1 + 3, 0, I) = "��茧"
        OutValue(1 + 4, 0, I) = "�{�錧"
        OutValue(1 + 5, 0, I) = "�H�c��"
        OutValue(1 + 6, 0, I) = "�R�`��"
        OutValue(1 + 7, 0, I) = "������"
        OutValue(1 + 8, 0, I) = "��錧"
        OutValue(1 + 9, 0, I) = "�Ȗ،�"
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
                    OutValue(I, InpCount + 1, 1) = InpArray(I)                  Rem 7���ԕ���
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
    For I = OutCount + 1 To UBound(DateList)
        OutDate = DateAdd("d", I - OutCount, DateList(OutCount))
        DateList(   I   ) = OutDate                         Rem ���t�ꗗ
        OutValue(0, I, 0) = OutDate                         Rem ����
        OutValue(0, I, 1) = OutDate                         Rem 7���ԕ���
        OutValue(0, I, 2) = OutDate                         Rem 10���l������
        OutValue(0, I, 3) = OutDate                         Rem 10���l������(7���ԕ���)
        OutValue(0, I, 4) = OutDate                         Rem 10���l������(�Z�o)
    Next
    OutCount = UBound(DateList)
    Rem -----------------------------------------------------------------------
    For I = 0 To 4
        OutFileName = "�����Ґ�." & I & ".txt"
        WScript.Echo "���o�F" & OutFileName
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            For J = 0 To OutCount - 1
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
        .Fields.Append "CD", 200, 128
        .Fields.Append "NAME", 200, 128
        .Fields.Append "VALUE", 5
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
    ReDim OutValue(17, UBound(DateList))

    OutValue(0, 0) = "���t"
    OutValue(1, 0) = "�����Ґ�"
    OutValue(2, 0) = "�����Ґ�(7���ԕ���)"
    OutValue(3, 0) = "�����Ґ�(10���l������)"
    OutValue(4, 0) = "�����Ґ�(10���l������E7���ԕ���)"
    OutValue(5, 0) = "�����Ґ�(�݌v)"
    OutValue(6, 0) = "���Ґ�"
    OutValue(7, 0) = "���Ґ�(7���ԕ���)"
    OutValue(8, 0) = "���Ґ�(�݌v)"
    OutValue(9, 0) = "�d�ǎҐ�"
    OutValue(10, 0) = "�d�ǎҐ�(7���ԕ���)"
    OutValue(11, 0) = "���@�×{��"
    OutValue(12, 0) = "�މ@�×{����"
    OutValue(13, 0) = "�މ@�×{����(�݌v)"
    OutValue(14, 0) = "PCR������(�����������)"
    OutValue(15, 0) = "PCR������(��������܂�)"
    OutValue(16, 0) = "�z����"
    OutValue(17, 0) = "�z����(7���ԕ��ϒl)"

    For I = 0 To OutCount - 1
        OutValue(0, I + 1) = DateList(I + 1)
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
                For I = 1 To OutCount - 1
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

            OutValue(5, InpCount + 1) = InpArray(1)         Rem �����Ґ�[�݌v�E�S��]
            InpCount = InpCount + 1
        Loop
Rem     OutCount = InpCount
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
                For I = 1 To OutCount - 1
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
            OutValue(8, InpCount + 1) = InpArray(1)         Rem ���Ґ�[�݌v�E�S��]
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
                For I = 1 To OutCount - 1
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
                For I = 1 To OutCount - 1
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
                For I = 1 To OutCount - 1
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            OutValue(14, InpCount + 1) = InpArray(7)        Rem PCR������(�����������)
            OutValue(15, InpCount + 1) = InpArray(9)        Rem PCR������(��������܂�)
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
        For I = 0 To OutCount - 1
            OutLine = ""
            For J = 0 To 15
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
    Call MakeExcelFile("�����Ґ�", "�����Ґ�.0.txt")
    Call MakeExcelFile("7���ԕ���", "�����Ґ�.1.txt")
    Call MakeExcelFile("10���l", "�����Ґ�.3.txt")
    Call MakeExcelFile("���{����", "���{����.txt")
    Call MakeExcelFile("���ʕt��", "���ʕt��.txt")

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("Sheet1")
Rem     .Visible = False
        .Activate
        .Name = "�O���t"
        Rem ---  1: �����Ґ� --------------------------------------------------
        Rem ---  2: 7���ԕ��� -------------------------------------------------
        Rem ---  3: 10���l������̊����Ґ� ------------------------------------
        With objDstWorkbook
            Call MakeGraph1(.Worksheets("�O���t"), .Worksheets("�����Ґ�"), .Worksheets("�O���t").ChartObjects.Add(0, 0, 912, 585), "�`��F1: �����Ґ�", "�����Ґ�", False)
            Call MakeGraph1(.Worksheets("�O���t"), .Worksheets("7���ԕ���"), .Worksheets("�O���t").ChartObjects.Add(960, 0, 912, 585), "�`��F2: 7���ԕ���", "�����Ґ���7���ԕ���", False)
            Call MakeGraph1(.Worksheets("�O���t"), .Worksheets("10���l"), .Worksheets("�O���t").ChartObjects.Add(0, 600, 912, 585), "�`��F3: 10���l������̊����Ґ�", "�����Ґ���10���l������7���ԕ���", False)
        End With
        Rem ---  4: �����Ґ��i�����s�j ----------------------------------------
        Erase Collection
        ReDim Collection(2, 4)
        Collection(0, 0) = "=""�����Ґ�"""
        Collection(0, 1) = "=�����Ґ�!$A$2:$A$"
        Collection(0, 2) = "=�����Ґ�!$O$2:$O$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""7���ԕ���"""
        Collection(1, 1) = "=�����Ґ�!$A$2:$A$"
        Collection(1, 2) = "=7���ԕ���!$O$2:$O$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""10���l"""
        Collection(2, 1) = "=�����Ґ�!$A$2:$A$"
        Collection(2, 2) = "=10���l!$O$2:$O$"
        Collection(2, 3) = 2
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("�O���t"), .Worksheets("�����Ґ�"), .Worksheets("�O���t").ChartObjects.Add(960, 600, 912, 585), "�`��F4: �����Ґ��i�����s�j", "�����Ґ��i�����s�j", False, Collection)
        End With
        Rem ---  5: �����Ґ��i���{�����j --------------------------------------
        Erase Collection
        ReDim Collection(2, 4)
        Collection(0, 0) = "=""�����Ґ�"""
        Collection(0, 1) = "=���{����!$A$2:$A$"
        Collection(0, 2) = "=���{����!$B$2:$B$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""7���ԕ���"""
        Collection(1, 1) = "=���{����!$A$2:$A$"
        Collection(1, 2) = "=���{����!$C$2:$C$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""10���l"""
        Collection(2, 1) = "=���{����!$A$2:$A$"
        Collection(2, 2) = "=���{����!$E$2:$E$"
        Collection(2, 3) = 2
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("�O���t"), .Worksheets("���{����"), .Worksheets("�O���t").ChartObjects.Add(0, 1200, 912, 585), "�`��F5: �����Ґ��i���{�����j", "�����Ґ��i���{�����j", False, Collection)
        End With
        Rem ---  6: ���ҏd�ǎҐ� ----------------------------------------------
        Erase Collection
        ReDim Collection(3, 4)
        Collection(0, 0) = "=""���Ґ�"""
        Collection(0, 1) = "=���{����!$A$2:$A$"
        Collection(0, 2) = "=���{����!$G$2:$G$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""���Ґ�(7���ԕ���)"""
        Collection(1, 1) = "=���{����!$A$2:$A$"
        Collection(1, 2) = "=���{����!$H$2:$H$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""�d�ǎҐ�"""
        Collection(2, 1) = "=���{����!$A$2:$A$"
        Collection(2, 2) = "=���{����!$J$2:$J$"
        Collection(2, 3) = 1
        Collection(3, 0) = "=""�d�ǎҐ�(7���ԕ���)"""
        Collection(3, 1) = "=���{����!$A$2:$A$"
        Collection(3, 2) = "=���{����!$K$2:$K$"
        Collection(3, 3) = 1
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("�O���t"), .Worksheets("���{����"), .Worksheets("�O���t").ChartObjects.Add(960, 1200, 912, 585), "�`��F6: ���ҏd�ǎҐ�", "���ҏd�ǎҐ�", False, Collection)
        End With
        Rem ---  7: ����7���Ԃ̐l��10���l������̐V�K�����Ґ� -----------------
        With objDstWorkbook
            MakeGraph1 .Worksheets("�O���t"), .Worksheets("10���l"), .Worksheets("�O���t").ChartObjects.Add(0, 1800, 912, 585), "�`��F7: ����7���Ԃ̐l��10���l������̐V�K�����Ґ�", "�����Ґ���10���l������7���ԕ��ρi�����j", True
            Rem --- �e�L�X�g�{�b�N�X�̕`�� ------------------------------------
            objExcel.Application.ScreenUpdating = False
            With .Worksheets("�O���t").ChartObjects(.Worksheets("�O���t").ChartObjects.Count).Chart
                With .Shapes.AddLabel(1, 0, 0, 72, 72)
                    With .TextFrame.Characters
                        .Text = "�O���t�͌����J���Ȃ̃f�[�^�[�A�ꗗ�\�͎Z�o�̂��߈�v���܂���"
                    End With
                    With .TextFrame2
                        .AutoSize = 1
                        .WordWrap = 0
                        With .TextRange.Font
                            .NameComplexScript = "Meiryo UI"
                            .NameFarEast = "Meiryo UI"
                            .Name = "Meiryo UI"
                            .Size = 8
                        End With
                    End With
                    .Fill.ForeColor.RGB = RGB(255, 255, 0)
                    .Top = 2
                    .Left = objDstWorkbook.Worksheets("�O���t").ChartObjects(objDstWorkbook.Worksheets("�O���t").ChartObjects.Count).Width - .Width - 12
                End With
                With .Shapes.AddLabel(1, 0, 0, 72, 72)
                    With .TextFrame.Characters
                        OutLine = ""
                        For I = 1 To 47
                            If RankData(4, I) <> "" Then
                                If OutLine = "" Then
                                    OutLine = RankData(4, I)
                                Else
                                    OutLine = OutLine & Chr(13) & Chr(10) & RankData(4, I)
                                End If
                            End If
                        Next
                        .Text = OutLine
                    End With
                    With .TextFrame2
                        .AutoSize = 1
                        .WordWrap = 0
                        .VerticalAnchor = 3
                        .HorizontalAnchor = 2
                        .TextRange.ParagraphFormat.LineRuleWithin = 0
                        .TextRange.ParagraphFormat.SpaceWithin = 9.6
                        .TextRange.ParagraphFormat.Alignment = 3
                        With .TextRange.Font
                            .NameComplexScript = "Meiryo UI"
                            .NameFarEast = "Meiryo UI"
                            .Name = "Meiryo UI"
                            .Size = 8
                        End With
                    End With
                    .Fill.ForeColor.RGB = RGB(226, 240, 217)
                    .Top = 26
                    .Left = objDstWorkbook.Worksheets("�O���t").ChartObjects(objDstWorkbook.Worksheets("�O���t").ChartObjects.Count).Width - .Width - 12
                End With
            End With
            objExcel.Application.ScreenUpdating = True
        End With
        Rem -------------------------------------------------------------------
Rem     .Visible = True
        .Activate
        Rem --- �O���t�ۑ� ----------------------------------------------------
        I = 0
        For Each objChart In .ChartObjects
            I = I + 1
            OutFileName = PicDir & "\pic" & I & "." & .Shapes("�O���t " & I).AlternativeText & ".png"
            WScript.Echo "�ۑ��F" & OutFileName
            .Range(objChart.TopLeftCell.Address(False, False)).Select
            objChart.Chart.Export OutFileName
        Next
    End With

    Rem -----------------------------------------------------------------------
    For Each objWorksheet In objDstWorkbook.Worksheets
        With objWorksheet
            .Activate
            If .Range("A1").Text = "���t" Then
                .Range("B" & (.Cells(.Rows.Count, 2).End(-4162).Row)).Select
            Else
                .Range("A1").Select
            End If
        End With
    Next

    Rem -----------------------------------------------------------------------
    OutFileName = CurDir & "\" & "covid-19.xlsx"
    WScript.Echo "�ۑ��F" & OutFileName
    objDstWorkbook.SaveAs (OutFileName)

    Rem -----------------------------------------------------------------------
    CopyExcel2Excel

    Rem --- ���ʕt���R�s�y�p --------------------------------------------------
    With objDstWorkbook.Worksheets("���ʕt��")
        RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
        .Activate
        .Range("E2:E" & RowsEnd).Copy
    End With

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("�O���t")
        .Activate
        .Range("A1").Select
    End With

    Rem -----------------------------------------------------------------------
    Set objDstWorkbook = Nothing
    Set objExcel = Nothing

Rem ---------------------------------------------------------------------------
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    WScript.Echo Now

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
Rem WScript.Quit

Rem ---------------------------------------------------------------------------
Sub MakeExcelFile(WorkSheetName, InpFileName)
    WScript.Echo "���o�F" & InpFileName
    objExcel.Application.ScreenUpdating = False
    objExcel.Workbooks.OpenText OutDir & "\" & InpFileName, 65001, , , , , True
    Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
    objSrcWorkbook.Worksheets(1).Name = WorkSheetName
    With objExcel
        objSrcWorkbook.Worksheets(1).Select
        .ActiveWindow.FreezePanes = False
        .Range("B2").Select
        .ActiveWindow.FreezePanes = True
        Select Case WorkSheetName
            Case "���ʕt��"
            Case Else
                For I = 2 To .Cells(1, 1).End(-4161).Column
                    With .Range(.Cells(2, I), .Cells(.Rows.Count, I).End(-4162))
                        .FormatConditions.AddTop10
                        With .FormatConditions(1)
                            .TopBottom = 1
                            .Rank = 1
                            .Percent = False
                            .Font.Color = -16776961
                            .Font.TintAndShade = 0
                            .StopIfTrue = False
                        End With
                    End With
                Next
        End Select
        Select Case WorkSheetName
            Case "�����Ґ�"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
            Case "7���ԕ���"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "10���l"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.0"
            Case "���{����"
                .Range("A1:P1").HorizontalAlignment = -4108
                .Range("A1:P" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
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
                .Range("P2:P" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
                .Range("G116").FormatConditions.Delete
            Case "���ʕt��"
                .Range("A1:E1").HorizontalAlignment = -4108
                .Range("A1:E" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
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
            .Range("B" & (.Cells(.Rows.Count, 2).End(-4162).Row)).Select
        Else
            .Range("B2").Select
        End If
    End With
    objSrcWorkbook.Worksheets(1).Move , objDstWorkbook.Worksheets(objDstWorkbook.Sheets.Count)
    Set objSrcWorkbook = Nothing
    objExcel.Application.ScreenUpdating = True
End Sub

Rem ---------------------------------------------------------------------------
Sub MakeGraph1(objGrphWorksheet, objSrcWorksheet, objChart, MessageText, ChartTitleText, LatestFlag)
    Dim I                               'For Next�p
    Dim RowsEnd                         '�f�[�^�[�̑��݂���ŏI�s
    Dim objRange1                       '�f�[�^�[�\�[�X�͈�
    Dim objRange2                       '�ő�l�̌����͈�
    Dim MaxValue                        '�ő�l�̒l
    Dim MaxColumn                       '�ő�l�̗�
    Dim MaxRow                          '�ő�l�̍s
    Dim LatestValue                     '�ŐV�̒l
    Dim LatestColumn                    '�ŐV�̗�
    Dim LatestRow                       '�ŐV�̍s
    Dim RecordsetCount                  '�f�[�^�[���x������
    Dim Point(1, 47, 3)                 '�f�[�^�[���x�����
    Dim objRecordset(1)                 '�\�[�g�p�I�u�W�F�N�g
    Dim objWorksheetRank                '
    Dim objRangeRank                    '

    WScript.Echo MessageText

    Rem --- ���ʕt���֘A ------------------------------------------------------
    Set objWorksheetRank = objDstWorkbook.Worksheets("���ʕt��")
    Set objRangeRank = objWorksheetRank.Range("B2:B4")
    Rem --- �f�[�^�[�\�[�X�֘A ------------------------------------------------
    With objSrcWorksheet
        LatestRow = .Cells(.Rows.Count, 2).End(-4162).Row   '���v��
        RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row     '���t��
        Set objRange1 = .Range("A1:AW" & RowsEnd)
        Set objRange2 = .Range("C2:AW" & RowsEnd)
    End With
    Rem --- �O���t�̕`�� ------------------------------------------------------
    With objGrphWorksheet
        With objChart.Chart
            .ChartArea.Font.Name = "Meiryo UI"
            .ChartArea.Font.Size = 8
            .HasTitle = True
            .ChartTitle.Text = ChartTitleText
            .ChartType = 4
            .Legend.Position = -4107
            .SetSourceData objRange1
            objExcel.Application.ScreenUpdating = True
            objGrphWorksheet.Range(objChart.TopLeftCell.Address(False, False)).Select
            objExcel.Application.ScreenUpdating = False
            Rem --- �O���t�̕`�� ----------------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 1 To .FullSeriesCollection.Count
                Select Case I
                    Case 2, 12, 13, 14, 15, 24, 28, 37, 41, 48
                        .FullSeriesCollection(I).IsFiltered = False
                    Case Else
                        .FullSeriesCollection(I).IsFiltered = True
                End Select
                If Not objRangeRank.Find((I - 1), , -4123, 1) Is Nothing Then
                    .FullSeriesCollection(I).IsFiltered = False
                End If
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �I�����ꂽ�O���t�̕`�� ------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 1 To .FullSeriesCollection.Count
                If .FullSeriesCollection(I).IsFiltered = False Then
                    Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 --------------
                    With objSrcWorksheet
                        Set objRange2 = .Range(.Cells(2, I + 1), .Cells(RowsEnd, I + 1))
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̕`�� (�ő�l) ---------------------
                    With .FullSeriesCollection(I).Points(MaxRow)
                        PointTop = .Top
                        .ApplyDataLabels
                        With .DataLabel
                            .ShowSeriesName = -1
                            .ShowCategoryName = -1
                            .ShowLegendKey = -0
                            .Separator = " "
                            .Font.Name = "Meiryo UI"
                            .Font.Size = 6
                            .Position = -4131
                            With .Format.TextFrame2
                                .AutoSize = 1
                                .WordWrap = 0
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .MarginBottom = 0
                            End With
                            With .Format.Fill
                                .Visible = -1
                                .ForeColor.RGB = RGB(255, 255, 0)
                            End With
                            .Left = .Left - 20
                            .Height = 8.5
                            .Top = PointTop - .Height / 2
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̕`�� (�ŐV�l) ---------------------
                    With .FullSeriesCollection(I).Points(LatestRow - 1)
                        PointTop = .Top
                        .ApplyDataLabels
                        With .DataLabel
                            .ShowSeriesName = -1
Rem                         .ShowCategoryName = -1
                            .ShowLegendKey = -0
                            .Separator = " "
                            .Font.Name = "Meiryo UI"
                            .Font.Size = 8
                            .Position = -4152
                            With .Format.TextFrame2
                                .AutoSize = 1
                                .WordWrap = 0
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .MarginBottom = 0
                            End With
                            With .Format.Fill
                                .Visible = -1
                                .ForeColor.RGB = RGB(255, 255, 0)
                            End With
                            .Left = .Left + 20
                            .Height = 8.5
                            .Top = PointTop - .Height / 2
                        End With
                    End With
                    RecordsetCount = RecordsetCount + 1
                End If
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �f�[�^�[���x���̎擾 --------------------------------------
            objExcel.Application.ScreenUpdating = False
            RecordsetCount = 0
            For I = 1 To .FullSeriesCollection.Count
                If .FullSeriesCollection(I).IsFiltered = False Then
                    Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 --------------
                    With objSrcWorksheet
                        Set objRange2 = .Range(.Cells(2, I + 1), .Cells(RowsEnd, I + 1))
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̎擾 (�ő�l) ---------------------
                    With .FullSeriesCollection(I).Points(MaxRow)
                        With .DataLabel
                            Point(0, RecordsetCount, 0) = I
                            Point(0, RecordsetCount, 1) = MaxRow
                            Point(0, RecordsetCount, 2) = .Top
                            Point(0, RecordsetCount, 3) = .Height
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̎擾 (�ŐV�l) ---------------------
                    With .FullSeriesCollection(I).Points(LatestRow - 1)
                        With .DataLabel
                            Point(1, RecordsetCount, 0) = I
                            Point(1, RecordsetCount, 1) = LatestRow - 1
                            Point(1, RecordsetCount, 2) = .Top
                            Point(1, RecordsetCount, 3) = .Height
                        End With
                    End With
                    RecordsetCount = RecordsetCount + 1
                End If
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �f�[�^�[���x���̒��� (������) -----------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To 1
                Set objRecordset(I) = CreateObject("ADODB.Recordset")
                With objRecordset(I)
                    .Fields.Append "CD", 5
                    .Fields.Append "POINT", 5
                    .Fields.Append "TOP", 5
                    .Fields.Append "HEIGHT", 5
                    .Open
                    Rem --- �f�[�^�[���x���̒��� (�擾) -----------------------
                    For J = 0 To RecordsetCount - 1
                        .AddNew
                        .Fields("CD").Value = Point(I, J, 0)
                        .Fields("POINT").Value = Point(I, J, 1)
                        .Fields("TOP").Value = Point(I, J, 2)
                        .Fields("HEIGHT").Value = Point(I, J, 3)
                        .Sort = "TOP DESC,CD"
                    Next
                    .MoveFirst
                    Rem --- �f�[�^�[���x���̒��� (�ݒ�) -----------------------
                    PosiTop = -1
                    For J = 1 To RecordsetCount
                        If PosiTop < 0 Then
                            PosiTop = .Fields("TOP").Value
                        ElseIf (PosiTop - .Fields("HEIGHT").Value) > .Fields("TOP").Value Then
                            PosiTop = .Fields("TOP").Value
                        Else
                            PosiTop = PosiTop - .Fields("HEIGHT").Value
                        End If
                        objChart.Chart.FullSeriesCollection(.Fields("CD").Value).Points(.Fields("POINT").Value).DataLabel.Top = PosiTop
                        .MoveNext
                    Next
                    .Close
                End With
                Set objRecordset(I) = Nothing
            Next
            objExcel.Application.ScreenUpdating = False
            Rem --- �\���͈͂̐ݒ� --------------------------------------------
            If LatestFlag = True Then
                .Axes(1, 1).MinimumScale = CDbl(CDate("2022/01/01"))
                .Axes(1, 1).MaximumScale = CDbl(CDate("2023/06/30"))
                .Axes(1, 1).MajorUnit = 7
                .Axes(1, 1).MajorUnitScale = 0
            Else
                .Axes(1, 1).MinimumScale = CDbl(CDate("2020/01/01"))
                .Axes(1, 1).MaximumScale = CDbl(CDate("2024/12/31"))
                .Axes(1, 1).MajorUnit = 1
                .Axes(1, 1).MajorUnitScale = 1
            End If
        End With
        .Shapes(.Shapes.Count).AlternativeText = ChartTitleText
    End With
    objExcel.Application.ScreenUpdating = True
End Sub

Rem ---------------------------------------------------------------------------
Sub MakeGraph2(objGrphWorksheet, objSrcWorksheet, objChart, MessageText, ChartTitleText, LatestFlag, Collection())
    Dim I                               'For Next�p
    Dim RowsEnd                         '�f�[�^�[�̑��݂���ŏI�s
    Dim objRange1                       '�f�[�^�[�\�[�X�͈�
    Dim objRange2                       '�ő�l�̌����͈�
    Dim MaxValue                        '�ő�l�̒l
    Dim MaxColumn                       '�ő�l�̗�
    Dim MaxRow                          '�ő�l�̍s
    Dim LatestValue                     '�ŐV�̒l
    Dim LatestColumn                    '�ŐV�̗�
    Dim LatestRow                       '�ŐV�̍s
    Dim RecordsetCount                  '�f�[�^�[���x������
    Dim Point(1, 47, 3)                 '�f�[�^�[���x�����
    Dim objRecordset(1)                 '�\�[�g�p�I�u�W�F�N�g
    Dim aryStrings                      '

    WScript.Echo MessageText
    Rem --- �f�[�^�[�\�[�X�֘A ------------------------------------------------
    With objSrcWorksheet
        LatestRow = .Cells(.Rows.Count, 2).End(-4162).Row   '���v��
        RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row     '���t��
        Set objRange1 = Nothing
        Set objRange2 = Nothing
    End With
    Rem --- �O���t�̕`�� ------------------------------------------------------
    With objGrphWorksheet
        With objChart.Chart
            .ChartArea.Font.Name = "Meiryo UI"
            .ChartArea.Font.Size = 8
            .HasTitle = True
            .ChartTitle.Text = ChartTitleText
            .ChartType = 4
            .Legend.Position = -4107
            objExcel.Application.ScreenUpdating = True
            objGrphWorksheet.Range(objChart.TopLeftCell.Address(False, False)).Select
            objExcel.Application.ScreenUpdating = False
            Rem --- �O���t�̕`�� ----------------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To UBound(Collection, 1)
                .SeriesCollection.NewSeries
                With .FullSeriesCollection(I + 1)
                    .Name = Collection(I, 0)
                    .XValues = Collection(I, 1) & RowsEnd
                    .Values = Collection(I, 2) & RowsEnd
                    .AxisGroup = Collection(I, 3)
                End With
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �I�����ꂽ�O���t�̕`�� ------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To .FullSeriesCollection.Count - 1
                If .FullSeriesCollection(I + 1).IsFiltered = False Then
                    Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 --------------
                    With .FullSeriesCollection(I + 1)
                        aryStrings = Split(Mid(Collection(I, 2), 2), "!")
                        If aryStrings(0) = "���{����" And Left(aryStrings(1), 2) = "$G" Then
                            Set objRange2 = objDstWorkbook.Worksheets(aryStrings(0)).Range("$G$117:$G$" & RowsEnd)
                        Else
                            Set objRange2 = objDstWorkbook.Worksheets(aryStrings(0)).Range(aryStrings(1) & RowsEnd)
                        End If
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̕`�� (�ő�l) ---------------------
                    With .FullSeriesCollection(I + 1).Points(MaxRow)
                        PointTop = .Top
                        .ApplyDataLabels
                        With .DataLabel
                            .ShowSeriesName = -1
                            .ShowCategoryName = -1
                            .ShowLegendKey = -0
                            .Separator = " "
                            .Font.Name = "Meiryo UI"
                            .Font.Size = 6
                            .Position = -4131
                            With .Format.TextFrame2
                                .AutoSize = 1
                                .WordWrap = 0
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .MarginBottom = 0
                            End With
                            With .Format.Fill
                                .Visible = -1
                                .ForeColor.RGB = RGB(255, 255, 0)
                            End With
                            .Left = .Left - 20
                            .Height = 8.5
                            .Top = PointTop - .Height / 2
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̕`�� (�ŐV�l) ---------------------
                    With .FullSeriesCollection(I + 1).Points(LatestRow - 1)
                        PointTop = .Top
                        .ApplyDataLabels
                        With .DataLabel
                            .ShowSeriesName = -1
Rem                         .ShowCategoryName = -1
                            .ShowLegendKey = -0
                            .Separator = " "
                            .Font.Name = "Meiryo UI"
                            .Font.Size = 8
                            .Position = -4152
                            With .Format.TextFrame2
                                .AutoSize = 1
                                .WordWrap = 0
                                .MarginLeft = 0
                                .MarginRight = 0
                                .MarginTop = 0
                                .MarginBottom = 0
                            End With
                            With .Format.Fill
                                .Visible = -1
                                .ForeColor.RGB = RGB(255, 255, 0)
                            End With
                            .Left = .Left + 20
                            .Height = 8.5
                            .Top = PointTop - .Height / 2
                        End With
                    End With
                    RecordsetCount = RecordsetCount + 1
                End If
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �f�[�^�[���x���̎擾 --------------------------------------
            objExcel.Application.ScreenUpdating = False
            RecordsetCount = 0
            For I = 0 To .FullSeriesCollection.Count - 1
                If .FullSeriesCollection(I + 1).IsFiltered = False Then
                    Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 --------------
                    With .FullSeriesCollection(I + 1)
                        aryStrings = Split(Mid(Collection(I, 2), 2), "!")
                        If aryStrings(0) = "���{����" And Left(aryStrings(1), 2) = "$G" Then
                            Set objRange2 = objDstWorkbook.Worksheets(aryStrings(0)).Range("$G$117:$G$" & RowsEnd)
                        Else
                            Set objRange2 = objDstWorkbook.Worksheets(aryStrings(0)).Range(aryStrings(1) & RowsEnd)
                        End If
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̎擾 (�ő�l) ---------------------
                    With .FullSeriesCollection(I + 1).Points(MaxRow)
                        With .DataLabel
                            Point(0, RecordsetCount, 0) = I + 1
                            Point(0, RecordsetCount, 1) = MaxRow
                            Point(0, RecordsetCount, 2) = .Top
                            Point(0, RecordsetCount, 3) = .Height
                        End With
                    End With
                    Rem --- �f�[�^�[���x���̎擾 (�ŐV�l) ---------------------
                    With .FullSeriesCollection(I + 1).Points(LatestRow - 1)
                        With .DataLabel
                            Point(1, RecordsetCount, 0) = I + 1
                            Point(1, RecordsetCount, 1) = LatestRow - 1
                            Point(1, RecordsetCount, 2) = .Top
                            Point(1, RecordsetCount, 3) = .Height
                        End With
                    End With
                    RecordsetCount = RecordsetCount + 1
                End If
            Next
            objExcel.Application.ScreenUpdating = True
            Rem --- �f�[�^�[���x���̒��� (������) -----------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To 1
                Set objRecordset(I) = CreateObject("ADODB.Recordset")
                With objRecordset(I)
                    .Fields.Append "CD", 5
                    .Fields.Append "POINT", 5
                    .Fields.Append "TOP", 5
                    .Fields.Append "HEIGHT", 5
                    .Open
                    Rem --- �f�[�^�[���x���̒��� (�擾) -----------------------
                    For J = 0 To RecordsetCount - 1
                        .AddNew
                        .Fields("CD").Value = Point(I, J, 0)
                        .Fields("POINT").Value = Point(I, J, 1)
                        .Fields("TOP").Value = Point(I, J, 2)
                        .Fields("HEIGHT").Value = Point(I, J, 3)
                        .Sort = "TOP DESC,CD"
                    Next
                    .MoveFirst
                    Rem --- �f�[�^�[���x���̒��� (�ݒ�) -----------------------
                    PosiTop = -1
                    For J = 1 To RecordsetCount
                        If PosiTop < 0 Then
                            PosiTop = .Fields("TOP").Value
                        ElseIf (PosiTop - .Fields("HEIGHT").Value) > .Fields("TOP").Value Then
                            PosiTop = .Fields("TOP").Value
                        Else
                            PosiTop = PosiTop - .Fields("HEIGHT").Value
                        End If
                        objChart.Chart.FullSeriesCollection(.Fields("CD").Value).Points(.Fields("POINT").Value).DataLabel.Top = PosiTop
                        .MoveNext
                    Next
                    .Close
                End With
                Set objRecordset(I) = Nothing
            Next
            objExcel.Application.ScreenUpdating = False
            Rem --- �\���͈͂̐ݒ� --------------------------------------------
            If LatestFlag = True Then
                .Axes(1, 1).MinimumScale = CDbl(CDate("2022/01/01"))
                .Axes(1, 1).MaximumScale = CDbl(CDate("2023/06/30"))
                .Axes(1, 1).MajorUnit = 7
                .Axes(1, 1).MajorUnitScale = 0
            Else
                .Axes(1, 1).MinimumScale = CDbl(CDate("2020/01/01"))
                .Axes(1, 1).MaximumScale = CDbl(CDate("2024/12/31"))
                .Axes(1, 1).MajorUnit = 1
                .Axes(1, 1).MajorUnitScale = 1
            End If
        End With
        .Shapes(.Shapes.Count).AlternativeText = ChartTitleText
    End With
    objExcel.Application.ScreenUpdating = True
End Sub

Rem ---------------------------------------------------------------------------
Sub CopyExcel2Excel()
    Set objOrgExcel = GetObject(, "Excel.Application")
    For Each objOrgWorkbook In objOrgExcel.Workbooks
        If objOrgWorkbook.Name = "_covid-19.xlsx" Then
            Ret = MsgBox("�f�[�^�[�̃R�s�[�����܂����H", vbYesNo)
            If Ret = 6 Then
                WScript.Echo "�]���FCovid19Data.xlsx��covid-19.xlsx"
                Rem --- �����Ґ� ----------------------------------------------
                With objDstWorkbook.Worksheets("�����Ґ�")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("�����Ґ�").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("�����Ґ�")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 7���ԕ��� ---------------------------------------------
                With objDstWorkbook.Worksheets("7���ԕ���")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("7���ԕ���").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("7���ԕ���")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 10���l ------------------------------------------------
                With objDstWorkbook.Worksheets("10���l")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("10���l").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("10���l")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- ���{���� ----------------------------------------------
                With objDstWorkbook.Worksheets("���{����")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("���{����").Range("A1:O" & RowsEnd).Value = .Range("A1:O" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("���{����")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- ���ʕt�� ----------------------------------------------
                With objDstWorkbook.Worksheets("���ʕt��")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("���ʕt��").Range("A1:E" & RowsEnd).Value = .Range("A1:E" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("���ʕt��")
                    .Activate
                    .Range("A1").Select
                End With
                Rem --- �O���t�p ----------------------------------------------
                With objDstWorkbook.Worksheets("���ʕt��")
                    RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
                    .Range("E2:E" & RowsEnd).Copy
                End With
                Rem --- �I������ ----------------------------------------------
                With objOrgWorkbook.Worksheets("�O���t")
                    .Activate
                    .Range("A1").Select
                End With
                objDstWorkbook.Close
                objExcel.Quit
            End If
            Exit For
        End If
    Next
    Set objOrgExcel = Nothing
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
