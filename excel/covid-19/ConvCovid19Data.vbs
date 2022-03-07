Rem ---------------------------------------------------------------------------
Rem ConvCovid19Data.vbs: CSV->TXT(TAB区切り)変換
Rem ---------------------------------------------------------------------------
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

    Erase OutData

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

    Erase OutData

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

    Erase OutData

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

Rem Erase OutData

    Wscript.Echo "終了：" & InpFileName

Rem --- nhk_news_covid19_prefectures_daily_data -------------------------------
    InpFileName="nhk_news_covid19_prefectures_daily_data.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutValue
    ReDim OutValue(49, 1999, 3)
    Rem 0: 各地の感染者数_1日ごとの発表数
    Rem 1: 各地の死者数_1日ごとの発表数
    Rem 2: 各地の直近1週間の人口10万人あたりの感染者数
    Rem 3: 各地の直近1週間の感染者数

    For I = 0 To 3
        OutValue(0, 0, I) = "日付"
        OutValue(1, 0, I) = "国内合計"
        OutValue(2, 0, I) = "空港検疫など"
    Next

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
            InpDate = InpArray(0)       Rem 日付
            InpCode = InpArray(1)       Rem 都道府県コード
            InpName = InpArray(2)       Rem 都道府県名
Rem         InpArray(3)                 Rem 各地の感染者数_1日ごとの発表数
Rem         InpArray(4)                 Rem 各地の感染者数_累計
Rem         InpArray(5)                 Rem 各地の死者数_1日ごとの発表数
Rem         InpArray(6)                 Rem 各地の死者数_累計
Rem         InpArray(7)                 Rem 各地の直近1週間の人口10万人あたりの感染者数
            If IsNumeric(InpCode) = True Then
                If OldCode <> InpCode Then
                    InpCount = 0
                    OldCode = InpCode
                    OutValue(InpCode + 2, 0, 0) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 1) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 2) = InpCode & ":" & InpName
                    OutValue(InpCode + 2, 0, 3) = InpCode & ":" & InpName
                End If
                If IsDate(OutValue(0, InpCount + 1, 0)) = False Then
                    OutValue(0, InpCount + 1, 0) = InpDate
                    OutValue(0, InpCount + 1, 1) = InpDate
                    OutValue(0, InpCount + 1, 2) = InpDate
                    OutValue(0, InpCount + 1, 3) = InpDate
                End If
                OutValue(InpCode + 2, InpCount + 1, 0) = InpArray(3)
                OutValue(InpCode + 2, InpCount + 1, 1) = InpArray(5)
                OutValue(InpCode + 2, InpCount + 1, 2) = InpArray(7)
            End If
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Wscript.Echo "計算：" & InpFileName
    For I = 0 To InpCount
        If OutData(0, I + 1) = OutValue(0, I + 1, 0) Then
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
                Next
            End If
        End If
    Next

    Wscript.Echo "書出：" & InpFileName
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

    Erase OutValue
    Erase OutData

    Wscript.Echo "終了：" & InpFileName

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
        Wscript.Echo "抽出：" & objFile.name
        objExcel.Workbooks.OpenText objFile.Path,65001,,,,,True
        Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
        Select Case objFile.name
            Case "nhk_news_covid19_domestic_daily_data.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "国内感染者"
            Case "nhk_news_covid19_prefectures_daily_data.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地感染者"
            Case "nhk_news_covid19_prefectures_daily_data.csv.1.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地死者数"
            Case "nhk_news_covid19_prefectures_daily_data.csv.2.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地10万人"
            Case "nhk_news_covid19_prefectures_daily_data.csv.3.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地 7日間"
            Case "pcr_case_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "PCR 検査数"
            Case "requiring_inpatient_care_etc_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "国内入退院"
            Case "severe_cases_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "国内重症者"
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

Rem ---------------------------------------------------------------------------
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
