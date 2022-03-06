Option Explicit

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

    Dim OutFileName
    Dim OutCount
    Dim OutLine
    Dim OutData()
    Dim OutValue()

    Dim OldCode
    Dim OldValue

    Dim TxtFileName
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
    Dim I, J
    Dim Ret

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

    Wscript.Echo "�I���F" & InpFileName

Rem --- nhk_news_covid19_prefectures_daily_data -------------------------------
    InpFileName="nhk_news_covid19_prefectures_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "�J�n�F" & InpFileName

    Erase OutValue
    ReDim OutValue(48, 1999)

    OutValue(0, 0) = "���t"
    OutValue(1, 0) = "��`���u�Ȃ�"

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
            InpDate = InpArray(0)
            InpCode = InpArray(1)
            InpName = InpArray(2)
            InpValue = InpArray(3)
            If IsNumeric(InpCode) = True Then
                If OldCode <> InpCode Then
                    InpCount = 0
                    OldCode = InpCode
                    OutValue(InpCode + 1, 0) = InpCode & ":" & InpName
                End If
                If IsDate(OutValue(0, InpCount + 1)) = False Then
                    OutValue(0, InpCount + 1) = InpDate
                End If
                OutValue(InpCode + 1, InpCount + 1) = InpValue
            End If
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "�v�Z�F" & InpFileName
    For I = 0 To InpCount
         InpValue = 0
         If OutData(0, I + 1) = OutValue(0, I + 1) Then
             For J = 0 To 47 - 1
                 InpValue = InpValue + OutValue(J + 2, I + 1)
             Next
             OutValue(1, I + 1) = OutData(1, I + 1) - InpValue
         End If
    Next

    Wscript.Echo "���o�F" & InpFileName
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            For J = 0 To 48
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

    Wscript.Echo "�I���F" & InpFileName

Rem ---------------------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

Rem ---------------------------------------------------------------------------
    Set objFolder = objFSO.GetFolder(OutDir)
    For Each objFile in objFolder.files
        Wscript.Echo "�J�n�F" & objFile
        objExcel.Workbooks.OpenText objFile.Path,65001,,,,,True
        Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
        For I = 1 To objSrcWorkbook.Sheets.Count
            Select Case objSrcWorkbook.WorkSheets(I).Name
                Case "nhk_news_covid19_domestic_daily"
                    objSrcWorkbook.WorkSheets(I).Name = "����������"
                Case "nhk_news_covid19_prefectures_da"
                    objSrcWorkbook.WorkSheets(I).Name = "�s���{����"
                Case "pcr_case_daily.csv"
                    objSrcWorkbook.WorkSheets(I).Name = "PCR ������"
                Case "requiring_inpatient_care_etc_da"
                    objSrcWorkbook.WorkSheets(I).Name = "���މ@�Ґ�"
                Case "severe_cases_daily.csv"
                    objSrcWorkbook.WorkSheets(I).Name = "�d�ǎҐ�"
            End Select
            objSrcWorkbook.WorkSheets(I).Move ,objDstWorkbook.WorkSheets(objDstWorkbook.Sheets.Count)
        Next
        Set objSrcWorkbook = Nothing
        Wscript.Echo "�I���F" & objFile
    Next

Rem TxtFileName = OutDir & "\" & "severe_cases_daily.csv" & ".txt"
Rem Wscript.Echo "�J�n�F" & TxtFileName
Rem objExcel.Workbooks.OpenText TxtFileName,65001,,,,,True                      Rem FileName,Origin,StartRow,DataType,TextQualifier,ConsecutiveDelimiter,Tab,Semicolon,Comma,Space,Other,OtherChar,FieldInfo,TextVisualLayout,DecimalSeparator,ThousandsSeparator,TrailingMinusNumbers,Local
Rem Wscript.Echo "�I���F" & TxtFileName

Rem ---------------------------------------------------------------------------
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
