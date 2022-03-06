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
    Wscript.Echo "開始：" & InpFileName

    Erase OutData
    ReDim OutData(1, 1999)

    OutData(0, 0) = "日付"
    OutData(1, 0) = "重症者数"

    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "書出：" & InpFileName
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

    Wscript.Echo "終了：" & InpFileName

Rem --- requiring_inpatient_care_etc_daily ------------------------------------
    InpFileName="requiring_inpatient_care_etc_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutData
    ReDim OutData(3, 1999)

    OutData(0, 0) = "日付"
    OutData(1, 0) = "入院者数"
    OutData(2, 0) = "退院者数"
    OutData(3, 0) = "確認予定"

    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "書出：" & InpFileName
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

    Wscript.Echo "終了：" & InpFileName

Rem --- pcr_case_daily --------------------------------------------------------
    InpFileName="pcr_case_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutData
    ReDim OutData(9, 1999)

    OutData(0, 0) = "日付"
    OutData(1, 0) = "国立感染症研究所"
    OutData(2, 0) = "検疫所"
    OutData(3, 0) = "地方衛生研究所・保健所"
    OutData(4, 0) = "民間検査会社（主に行政検査）"
    OutData(5, 0) = "大学等"
    OutData(6, 0) = "医療機関"
    OutData(7, 0) = "小計"
    OutData(8, 0) = "民間検査会社（主に自費検査）"
    OutData(9, 0) = "計"

    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "書出：" & InpFileName
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

    Wscript.Echo "終了：" & InpFileName

Rem --- nhk_news_covid19_domestic_daily_data ----------------------------------
    InpFileName="nhk_news_covid19_domestic_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutData
    ReDim OutData(4, 1999)

    OutData(0, 0) = "日付"
    OutData(1, 0) = "国内感染者数"
    OutData(2, 0) = "国内感染者累計"
    OutData(3, 0) = "国内死者者数"
    OutData(4, 0) = "国内死者数累計"

    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "書出：" & InpFileName
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

    Wscript.Echo "終了：" & InpFileName

Rem --- nhk_news_covid19_prefectures_daily_data -------------------------------
    InpFileName="nhk_news_covid19_prefectures_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutValue
    ReDim OutValue(48, 1999)

    OutValue(0, 0) = "日付"
    OutValue(1, 0) = "空港検疫など"

    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "計算：" & InpFileName
    For I = 0 To InpCount
         InpValue = 0
         If OutData(0, I + 1) = OutValue(0, I + 1) Then
             For J = 0 To 47 - 1
                 InpValue = InpValue + OutValue(J + 2, I + 1)
             Next
             OutValue(1, I + 1) = OutData(1, I + 1) - InpValue
         End If
    Next

    Wscript.Echo "書出：" & InpFileName
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

    Wscript.Echo "終了：" & InpFileName

Rem ---------------------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

Rem ---------------------------------------------------------------------------
    Set objFolder = objFSO.GetFolder(OutDir)
    For Each objFile in objFolder.files
        Wscript.Echo "開始：" & objFile
        objExcel.Workbooks.OpenText objFile.Path,65001,,,,,True
        Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
        For I = 1 To objSrcWorkbook.Sheets.Count
            Select Case objSrcWorkbook.WorkSheets(I).Name
                Case "nhk_news_covid19_domestic_daily"
                    objSrcWorkbook.WorkSheets(I).Name = "国内感染者"
                Case "nhk_news_covid19_prefectures_da"
                    objSrcWorkbook.WorkSheets(I).Name = "都道府県別"
                Case "pcr_case_daily.csv"
                    objSrcWorkbook.WorkSheets(I).Name = "PCR 検査数"
                Case "requiring_inpatient_care_etc_da"
                    objSrcWorkbook.WorkSheets(I).Name = "入退院者数"
                Case "severe_cases_daily.csv"
                    objSrcWorkbook.WorkSheets(I).Name = "重症者数"
            End Select
            objSrcWorkbook.WorkSheets(I).Move ,objDstWorkbook.WorkSheets(objDstWorkbook.Sheets.Count)
        Next
        Set objSrcWorkbook = Nothing
        Wscript.Echo "終了：" & objFile
    Next

Rem TxtFileName = OutDir & "\" & "severe_cases_daily.csv" & ".txt"
Rem Wscript.Echo "開始：" & TxtFileName
Rem objExcel.Workbooks.OpenText TxtFileName,65001,,,,,True                      Rem FileName,Origin,StartRow,DataType,TextQualifier,ConsecutiveDelimiter,Tab,Semicolon,Comma,Space,Other,OtherChar,FieldInfo,TextVisualLayout,DecimalSeparator,ThousandsSeparator,TrailingMinusNumbers,Local
Rem Wscript.Echo "終了：" & TxtFileName

Rem ---------------------------------------------------------------------------
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
