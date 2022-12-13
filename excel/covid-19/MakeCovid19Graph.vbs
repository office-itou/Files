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
    Dim StartTime
    Dim EndTime

    StartTime = CDate(Now)
    WScript.Echo FormatDateTime(StartTime)

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
    Dim objRangeMax
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
Rem Dim Collection()

Rem ---------------------------------------------------------------------------
Class ClassCollection
    Public Name                                             '
    Public XValues                                          '
    Public Values                                           '
    Public AxisGroup                                        '
End Class

Class ClassGraph
    Private I

    Public WorksheetGrph                                    'Worksheet:�O���t
    Public WorksheetData                                    'Worksheet:�f�[�^�[
    Public Left                                             '�`��͈�:������̈ʒu
    Public Top                                              '   �V   :�ォ��̈ʒu
    Public Width                                            '   �V   :�`�悷�镝
    Public Height                                           '   �V   :�`�悷�鍂��
    Public ChartTitleText                                   '�O���t�^�C�g��
    Public Collection(3)                                    '�O���t���

    Private Sub Class_Initialize()
        For I = LBound(Collection) To UBound(Collection)
            Set Collection(I) = New ClassCollection
        Next
    End Sub

    Private Sub Class_Terminate()
        For I = LBound(Collection) To UBound(Collection)
            Set Collection(I) = Nothing
        Next
    End Sub
End Class

    Dim clsGraph

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
                    DateList(InpCount + 1) = InpArray(I)                        '���t�ꗗ
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  '����
                    OutValue(I, InpCount + 1, 1) = InpArray(I)                  '7���ԕ���
                    OutValue(I, InpCount + 1, 2) = InpArray(I)                  '10���l������
                    OutValue(I, InpCount + 1, 3) = InpArray(I)                  '10���l������(7���ԕ���)
                    OutValue(I, InpCount + 1, 4) = InpArray(I)                  '10���l������(�Z�o)
                Else
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  '����
                    If InpCount >= 6 Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 0)
                        Next
                        OutValue(I, InpCount + 1, 1) = Round(InpValue / 7, 2)   '7���ԕ���
                                                                                '10���l������(�Z�o)
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
                OutValue(I, InpCount + 1, 2) = InpArray(I)                      '10���l������(����)
                If InpCount >= 6 Then
                    If IsNumeric(OutValue(I, InpCount + 1 - 7, 2)) = True Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 2)
                        Next
                        OutValue(I, InpCount + 1, 3) = InpValue                 '10���l������(7���ԍ��v)
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
        DateList(I) = OutDate                               '���t�ꗗ
        OutValue(0, I, 0) = OutDate                         '����
        OutValue(0, I, 1) = OutDate                         '7���ԕ���
        OutValue(0, I, 2) = OutDate                         '10���l������
        OutValue(0, I, 3) = OutDate                         '10���l������(7���ԕ���)
        OutValue(0, I, 4) = OutDate                         '10���l������(�Z�o)
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
            .Fields("CD").Value = I + 1                                         '�s���{���R�[�h
            .Fields("NAME").Value = OutValue(I + 2, 0, 4)                       '�s���{����
            .Fields("VALUE").Value = OutValue(I + 2, InpCount + 0, 4)           '�e�n�̒���1�T�Ԃ̐l��10���l������̊����Ґ�
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
                                                            '�����Ґ�[�����E�S��]
            If IsNumeric(OutValue(4, InpCount + 1 - 1)) = False Then
                OutValue(1, InpCount + 1) = InpArray(1)
            Else
                OutValue(1, InpCount + 1) = InpArray(1) - OutValue(5, InpCount + 1 - 1)
            End If
                                                            '10���l������(����)
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
                OutValue(2, InpCount + 1) = Round(InpValue / 7, 2)              '7���ԕ���
                If IsNumeric(OutValue(3, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For J = 0 To 6
                        InpValue = InpValue + OutValue(3, InpCount + 1 - J)
                    Next
                    OutValue(4, InpCount + 1) = InpValue                        '10���l������(7���ԍ��v)
                End If
            End If

            OutValue(5, InpCount + 1) = InpArray(1)         '�����Ґ�[�݌v�E�S��]
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
                                                            '���Ґ�[���ʁE�S��]
            If IsNumeric(OutValue(6, InpCount + 1 - 1)) = False Then
                OutValue(6, InpCount + 1) = InpArray(1)
            Else
                OutValue(6, InpCount + 1) = InpArray(1) - OutData
            End If
            OutData = InpArray(1)
                                                            '7���ԕ���
            If InpCount >= 6 Then
                If IsNumeric(OutValue(6, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For I = 0 To 6
                        InpValue = InpValue + OutValue(6, InpCount + 1 - I)
                    Next
                    OutValue(7, InpCount + 1) = Round(InpValue / 7, 2)
                End If
            End If
            OutValue(8, InpCount + 1) = InpArray(1)         '���Ґ�[�݌v�E�S��]
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
            OutValue(9, InpCount + 1) = InpArray(1)         '�d�ǎҐ�[���ʁE�S��]
                                                            '7���ԕ���
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
            OutValue(11, InpCount + 1) = InpArray(1)        '���@�×{��[���ʁE�S��]
                                                            '�މ@�×{����[���ʁE�S��]
            If IsNumeric(OutValue(12, InpCount + 1 - 1)) = False Then
                OutValue(12, InpCount + 1) = InpArray(2)
            Else
                OutValue(12, InpCount + 1) = InpArray(2) - OutData
            End If
            OutValue(13, InpCount + 1) = InpArray(2)        '�މ@�×{����[�݌v�E�S��]
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
            OutValue(14, InpCount + 1) = InpArray(7)        'PCR������(�����������)
            OutValue(15, InpCount + 1) = InpArray(9)        'PCR������(��������܂�)
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
    End With
    Rem ---  1: �����Ґ� ------------------------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("�����Ґ�")
        .Left = 0
        .Top = 0
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ�"
        Call MakeGraph(clsGraph, "�`��F1: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  2: 7���ԕ��� -----------------------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("7���ԕ���")
        .Left = 960
        .Top = 0
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ���7���ԕ���"
        Call MakeGraph(clsGraph, "�`��F2: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  3: 10���l������̊����Ґ� ----------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("10���l")
        .Left = 0
        .Top = 600
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ���10���l������7���ԕ���"
        Call MakeGraph(clsGraph, "�`��F3: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  4: �����Ґ��i�����s�j --------------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("�����Ґ�")
        .Left = 960
        .Top = 600
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ��i�����s�j"
        With .Collection(0)
            .Name = "=""�����Ґ�"""
            .XValues = "=�����Ґ�!$A$2:$A$"
            .Values = "=�����Ґ�!$O$2:$O$"
            .AxisGroup = 1
        End With
        With .Collection(1)
            .Name = "=""7���ԕ���"""
            .XValues = "=�����Ґ�!$A$2:$A$"
            .Values = "=7���ԕ���!$O$2:$O$"
            .AxisGroup = 1
        End With
        With .Collection(2)
            .Name = "=""10���l"""
            .XValues = "=�����Ґ�!$A$2:$A$"
            .Values = "=10���l!$O$2:$O$"
            .AxisGroup = 2
        End With
        Call MakeGraph(clsGraph, "�`��F4: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  5: �����Ґ��i���{�����j ------------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("���{����")
        .Left = 0
        .Top = 1200
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ��i���{�����j"
        With .Collection(0)
            .Name = "=""�����Ґ�"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$B$2:$B$"
            .AxisGroup = 1
        End With
        With .Collection(1)
            .Name = "=""7���ԕ���"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$C$2:$C$"
            .AxisGroup = 1
        End With
        With .Collection(2)
            .Name = "=""10���l"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$E$2:$E$"
            .AxisGroup = 2
        End With
        Call MakeGraph(clsGraph, "�`��F5: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  6: ���ҏd�ǎҐ� --------------------------------------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("���{����")
        .Left = 960
        .Top = 1200
        .Width = 912
        .Height = 585
        .ChartTitleText = "���ҏd�ǎҐ��i���{�����j"
        With .Collection(0)
            .Name = "=""���Ґ�"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$G$2:$G$"
            .AxisGroup = 1
        End With
        With .Collection(1)
            .Name = "=""���Ґ�(7���ԕ���)"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$H$2:$H$"
            .AxisGroup = 1
        End With
        With .Collection(2)
            .Name = "=""�d�ǎҐ�"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$J$2:$J$"
            .AxisGroup = 1
        End With
        With .Collection(3)
            .Name = "=""�d�ǎҐ�(7���ԕ���)"""
            .XValues = "=���{����!$A$2:$A$"
            .Values = "=���{����!$K$2:$K$"
            .AxisGroup = 1
        End With
        Call MakeGraph(clsGraph, "�`��F6: " & .ChartTitleText, False)
    End With
    Set clsGraph = Nothing
    Rem ---  7: ����7���Ԃ̐l��10���l������̐V�K�����Ґ� ---------------------
    Set clsGraph = New ClassGraph
    With clsGraph
        Set .WorksheetGrph = objDstWorkbook.Worksheets("�O���t")
        Set .WorksheetData = objDstWorkbook.Worksheets("10���l")
        .Left = 0
        .Top = 1800
        .Width = 912
        .Height = 585
        .ChartTitleText = "�����Ґ���10���l������7���ԕ��ρi�����j"
        Call MakeGraph(clsGraph, "�`��F7: " & .ChartTitleText, True)
    End With
    Set clsGraph = Nothing
    Rem --- �e�L�X�g�{�b�N�X�̕`�� --------------------------------------------
    objExcel.Application.ScreenUpdating = False
    With objDstWorkbook.Worksheets("�O���t").ChartObjects(objDstWorkbook.Worksheets("�O���t").ChartObjects.Count).Chart
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
                    .Size = 6
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
    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("�O���t")
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
    EndTime = CDate(Now)

    WScript.Echo "�o�߁F" & FormatSecond2DateTime(DateDiff("s", StartTime, EndTime))
    WScript.Echo FormatDateTime(EndTime)

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
Sub MakeGraph(clsGraph, MessageText, LatestFlag)
    Dim I                               'For Next�p
    Dim objWorksheetGrph                'Worksheet:�O���t
    Dim objWorksheetData                'Worksheet:�f�[�^�[
    Dim objChart                        'Chart
    Dim RowsEnd                         '�f�[�^�[�̑��݂���ŏI�s
    Dim objRangeMax                     '�ő�l�̌����͈�
    Dim MaxValue                        '�ő�l�̒l
    Dim MaxColumn                       '�ő�l�̗�
    Dim MaxRow                          '�ő�l�̍s
    Dim LatestValue                     '�ŐV�̒l
    Dim LatestColumn                    '�ŐV�̗�
    Dim LatestRow                       '�ŐV�̍s
    Dim RecordsetCount                  '�f�[�^�[���x������
    Dim Point(1, 47, 3)                 '�f�[�^�[���x�����
    Dim objRecordset(1)                 '�\�[�g�p�I�u�W�F�N�g
    Dim objWorksheetRank                'Worksheet:���ʕt��
    Dim objRangeRank                    '���ʕt����ʂR�ʂ܂ł͈̔�
    Dim aryStrings                      '

    WScript.Echo MessageText
    Rem -----------------------------------------------------------------------
    Set objWorksheetGrph = clsGraph.WorksheetGrph
    Set objWorksheetData = clsGraph.WorksheetData
    With clsGraph
        Set objChart = .WorksheetGrph.ChartObjects.Add(.Left, .Top, .Width, .Height)
    End With
    Rem -----------------------------------------------------------------------
    objExcel.Application.ScreenUpdating = True
    objWorksheetGrph.Range(objChart.TopLeftCell.Address(False, False)).Select
    objExcel.Application.ScreenUpdating = False
    Rem --- ���ʕt���֘A ------------------------------------------------------
    Set objWorksheetRank = objDstWorkbook.Worksheets("���ʕt��")
    Set objRangeRank = objWorksheetRank.Range("B2:B4")
    Rem --- �f�[�^�[�\�[�X�֘A ------------------------------------------------
    With objWorksheetData
        LatestRow = .Cells(.Rows.Count, 2).End(-4162).Row   '���v��
        RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row     '���t��
    End With
    Rem --- �O���t�̕`�� ------------------------------------------------------
    With objChart.Chart
        .ChartArea.Font.Name = "Meiryo UI"
        .ChartArea.Font.Size = 8
        Rem --- �^�C�g�� ------------------------------------------------------
        .HasTitle = True
        .ChartTitle.Text = clsGraph.ChartTitleText
        .ChartTitle.Font.Size = 8
        Rem --- �O���t�̎�� --------------------------------------------------
        .ChartType = 4
        Rem --- �}�� ----------------------------------------------------------
        .Legend.Position = -4107
        .Legend.Font.Size = 6
        Rem --- �f�[�^�[�̑I�� ------------------------------------------------
        If clsGraph.Collection(LBound(clsGraph.Collection)).Name = "" Then
            .SetSourceData objWorksheetData.Range("A1:AW" & RowsEnd)
        Else
            For I = LBound(clsGraph.Collection) To UBound(clsGraph.Collection)
                If clsGraph.Collection(I).Name <> "" Then
                    .SeriesCollection.NewSeries
                    With .FullSeriesCollection(I + 1)
                        .Name = clsGraph.Collection(I).Name
                        .XValues = clsGraph.Collection(I).XValues & RowsEnd
                        .Values = clsGraph.Collection(I).Values & RowsEnd
                        .AxisGroup = clsGraph.Collection(I).AxisGroup
                    End With
                End If
            Next
        End If
        Rem --- �\���͈͂̐ݒ� ------------------------------------------------
        If LatestFlag = True Then
            .Axes(1, 1).MinimumScale = CDbl(CDate("2022/01/01"))
            .Axes(1, 1).MaximumScale = CDbl(CDate("2023/06/30"))
            .Axes(1, 1).MajorUnit = 7                       '7���P��
            .Axes(1, 1).MajorUnitScale = 0                  '���P��
        Else
            .Axes(1, 1).MinimumScale = CDbl(CDate("2020/01/01"))
            .Axes(1, 1).MaximumScale = CDbl(CDate("2024/12/31"))
            .Axes(1, 1).MajorUnit = 1                       '1���P��
            .Axes(1, 1).MajorUnitScale = 1                  '���P��
        End If
        Rem --- �c�����̐ݒ� --------------------------------------------------
        .Axes(1, 1).TickLabels.Font.Size = 6                '���i���ځj��
        .Axes(2, 1).TickLabels.Font.Size = 6                '�c�i�l�j���i�v���C�}���[�j
        If .Axes.Count > 2 Then
            .Axes(2, 2).TickLabels.Font.Size = 6            '�c�i�l�j���i�Z�J���_���[�j
        End If
        objExcel.Application.ScreenUpdating = True
        Rem -------------------------------------------------------------------
        objExcel.Application.ScreenUpdating = False
        For I = 1 To .FullSeriesCollection.Count
            If clsGraph.Collection(LBound(clsGraph.Collection)).Name = "" Then
                Rem --- �O���t�̕\������ --------------------------------------
                Select Case I
                    Case 1
                        .FullSeriesCollection(I).IsFiltered = True
                    Case 2, 12, 13, 14, 15, 24, 28, 37, 41, 47, 48
                        .FullSeriesCollection(I).IsFiltered = False
                    Case Else
                        .FullSeriesCollection(I).IsFiltered = True
                End Select
                If Not objRangeRank.Find((I - 1), , -4123, 1) Is Nothing Then
                    .FullSeriesCollection(I).IsFiltered = False
                End If
                Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 ------------------
                If .FullSeriesCollection(I).IsFiltered = False Then
                    With objWorksheetData
                        Set objRangeMax = .Range(.Cells(2, I + 1), .Cells(LatestRow, I + 1))
                        MaxValue = objExcel.Max(objRangeMax)
                        With objRangeMax.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                End If
            Else
                Rem --- �I�����ꂽ�O���t�̍ő�l�̈ʒu���擾 ------------------
                If .FullSeriesCollection(I).IsFiltered = False Then
                    aryStrings = Split(Mid(clsGraph.Collection(I - 1).Values, 2), "!")
                    If aryStrings(0) = "���{����" And Left(aryStrings(1), 2) = "$G" Then
                        Set objRangeMax = objDstWorkbook.Worksheets(aryStrings(0)).Range("$G$117:$G$" & LatestRow)
                    Else
                        Set objRangeMax = objDstWorkbook.Worksheets(aryStrings(0)).Range(aryStrings(1) & LatestRow)
                    End If
                    MaxValue = objExcel.Max(objRangeMax)
                    With objRangeMax.Find(MaxValue, , -4123, 1)
                        MaxColumn = .Column - 1
                        MaxRow = .Row - 1
                    End With
                End If
            End If
            Rem --- �f�[�^�[���x���̕`�� (�ő�l) -----------------------------
            If .FullSeriesCollection(I).IsFiltered = False Then
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
                        .Height = .Font.Size + 1.5
                        .Top = PointTop - .Height / 2
                    End With
                End With
                Rem --- �f�[�^�[���x���̕`�� (�ŐV�l) -------------------------
                With .FullSeriesCollection(I).Points(LatestRow - 1)
                    PointTop = .Top
                    .ApplyDataLabels
                    With .DataLabel
                        .ShowSeriesName = -1
Rem                     .ShowCategoryName = -1
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
                        .Height = .Font.Size + 1.5
                        .Top = PointTop - .Height / 2
                    End With
                End With
                Rem --- �f�[�^�[���x���̎擾 (�ő�l) -------------------------
                Point(0, RecordsetCount, 0) = I
                Point(0, RecordsetCount, 1) = MaxRow
                Rem --- �f�[�^�[���x���̎擾 (�ŐV�l) -------------------------
                Point(1, RecordsetCount, 0) = I
                Point(1, RecordsetCount, 1) = LatestRow - 1
                RecordsetCount = RecordsetCount + 1
            End If
        Next
        Rem --- �f�[�^�[���x���̎擾 ------------------------------------------
        objExcel.Application.ScreenUpdating = True
        For I = 0 To 1
            For J = 0 To RecordsetCount - 1
                With .FullSeriesCollection(Point(I, J, 0)).Points(Point(I, J, 1)).DataLabel
                    Point(I, J, 2) = .Top
                    Point(I, J, 3) = .Height
                End With
            Next
        Next
        objExcel.Application.ScreenUpdating = False
        Rem --- �f�[�^�[���x���̒��� ------------------------------------------
        For I = 0 To 1
            Set objRecordset(I) = CreateObject("ADODB.Recordset")
            With objRecordset(I)
                Rem --- �f�[�^�[���x���̒��� (������) -------------------------
                .Fields.Append "CD", 5
                .Fields.Append "POINT", 5
                .Fields.Append "TOP", 5
                .Fields.Append "HEIGHT", 5
                .Open
                Rem --- �f�[�^�[���x���̒��� (�擾) ---------------------------
                For J = 0 To RecordsetCount - 1
                    .AddNew
                    .Fields("CD").Value = Point(I, J, 0)
                    .Fields("POINT").Value = Point(I, J, 1)
                    .Fields("TOP").Value = Point(I, J, 2)
                    .Fields("HEIGHT").Value = Point(I, J, 3)
                    .Sort = "TOP DESC,CD"
                Next
                .MoveFirst
                Rem --- �f�[�^�[���x���̒��� (�ݒ�) ---------------------------
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
    End With
    With objWorksheetGrph
        .Shapes(.Shapes.Count).AlternativeText = clsGraph.ChartTitleText
    End With
    objExcel.Application.ScreenUpdating = True
End Sub

Function FormatDateTime(DateTimeValue)
    Dim strYear
    Dim strMonth
    Dim strDay
    Dim strHour
    Dim strMinute
    Dim strSecond
    Dim strWeekday

    If IsDate(DateTimeValue) = False Then
        FormatDateTime = Null
        Exit Function
    End If

    strYear = Right("0000" & Year(Now), 4)
    strMonth = Right("00" & Month(DateTimeValue), 2)
    strDay = Right("00" & Day(DateTimeValue), 2)
    strHour = Right("00" & Hour(DateTimeValue), 2)
    strMinute = Right("00" & Minute(DateTimeValue), 2)
    strSecond = Right("00" & Second(DateTimeValue), 2)
    strWeekday = WeekdayName(Weekday(DateTimeValue), True)

    FormatDateTime = strYear & "/" & strMonth & "/" & strDay & "(" & strWeekday & ")" & " " & strHour & ":" & strMinute & ":" & strSecond
End Function

Function FormatSecond2DateTime(SecondValue)
    Dim strHour
    Dim strMinute
    Dim strSecond

    If (SecondValue \ 3600) < 100 Then
        strHour = Right("00" & SecondValue \ 3600, 2)
    Else
        strHour = SecondValue \ 3600
    End If
    strMinute = Right("00" & SecondValue \ 60 Mod 60, 2)
    strSecond = Right("00" & SecondValue Mod 60, 2)

    FormatSecond2DateTime = strHour & ":" & strMinute & ":" & strSecond
End Function

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
