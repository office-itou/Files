Option Explicit

    Dim InpDir
    Dim InpFileName
    Dim InpCount
    Dim InpLine
    Dim InpArray
    Dim InpValue

    Dim OutDir
    Dim OutFileName
    Dim OutCount
    Dim OutLine
    Dim OutData()
    Dim OutValue()

    Dim Ret
    Dim I
    Dim J
    Dim K

Rem With Application.FileDialog(msoFileDialogOpen)
Rem     .AllowMultiSelect = False
Rem     .Show
Rem     InpFileName = .SelectedItems(1)
Rem End With

    InpDir = ".\data"
    OutDir = ".\conv"

Rem --- nhk_news_covid19_domestic_daily_data.csv ------------------------------
    Erase OutData
    ReDim OutData(5, 1500)

    InpFileName = "nhk_news_covid19_domestic_daily_data.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount < 0 Then
                OutData(0, 0) = "日付"
                OutData(1, 0) = "国内感染者数"
                OutData(2, 0) = "国内感染者累計"
                OutData(3, 0) = "国内死者者数"
                OutData(4, 0) = "国内死者数累計"
                InpCount = 1
            Else
                OutData(0, InpCount) = InpArray(0)
                OutData(1, InpCount) = InpArray(1)
                OutData(2, InpCount) = InpArray(2)
                OutData(3, InpCount) = InpArray(3)
                OutData(4, InpCount) = InpArray(4)
                InpCount = InpCount + 1
            End If
Rem         DoEvents
        Loop
        .Close
    End With

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount - 1
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I) & Chr(9) & OutData(4, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

    OutCount = InpCount - 1

Rem --- nhk_news_covid19_prefectures_daily_data.csv ---------------------------
    Erase OutValue
    ReDim OutValue(48, 1500)
    Dim MaxCount
    Dim InpDate
    Dim InpCode
    Dim InpName
    Dim OldCode

    InpFileName = "nhk_news_covid19_prefectures_daily_data.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCode = 0
        OldCode = -1
        InpCount = 0
        MaxCount = 0
        OutValue(0, 0) = "日付"
        OutValue(1, 0) = "空港検疫など"
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            InpDate = InpArray(0)
            InpCode = InpArray(1)
            InpName = InpArray(2)
            InpValue = InpArray(3)
            If IsNumeric(InpCode) = True Then
                If OldCode <> InpCode Then
                    OutValue(InpCode + 1, 0) = InpCode & ":" & InpName
                    InpCount = 1
                    OldCode = InpCode
                End If
                If OutValue(0, InpCount) = "" Then
                    OutValue(          0, InpCount) = InpDate
                    OutValue(InpCode + 1, InpCount) = InpValue
                    MaxCount = MaxCount + 1
                Else
                    For I = 1 To MaxCount
                        If OutValue(0, I) = InpDate Then
                            OutValue(InpCode + 1, InpCount) = InpValue
                            Exit For
                        End If
                    Next
                End If
                InpCount = InpCount + 1
            End If
Rem         DoEvents
        Loop
        .Close
    End With

    For I = 1 To MaxCount
        For J = 1 To UBound(OutData, 2)
            If OutValue(0, I) = OutData(0, J) Then
                InpValue = 0
                For K = 2 To 48
                    InpValue = InpValue + OutValue(K, I)
                Next
                OutValue(1, I) = OutData(1, J) - InpValue
                Exit For
            End If
        Next
    Next

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To MaxCount
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

Rem --- severe_cases_daily.csv ------------------------------------------------
    Erase OutData
    ReDim OutData(1, 1500)

    InpFileName = "severe_cases_daily.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount < 0 Then
                OutData(0, 0) = "日付"
                OutData(1, 0) = "重症者数"
                InpCount = 1
            Else
                OutData(0, InpCount) = InpArray(0)
                OutData(1, InpCount) = InpArray(1)
                InpCount = InpCount + 1
            End If
Rem         DoEvents
        Loop
        .Close
    End With

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount - 1
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

Rem --- requiring_inpatient_care_etc_daily.csv --------------------------------
    Erase OutData
    ReDim OutData(3, 1500)

    InpFileName = "requiring_inpatient_care_etc_daily.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount < 0 Then
                OutData(0, 0) = "日付"
                OutData(1, 0) = "入院者数"
                OutData(2, 0) = "退院者数"
                OutData(3, 0) = "確認予定"
                InpCount = 1
            Else
                OutData(0, InpCount) = InpArray(0)
                OutData(1, InpCount) = InpArray(1)
                If InpCount = 1 Then
                    InpValue = InpArray(2)
                Else
                    OutData(2, InpCount) = InpArray(2) - InpValue
                    InpValue = InpArray(2)
                End If
                OutData(3, InpCount) = InpArray(3)
                InpCount = InpCount + 1
            End If
Rem         DoEvents
        Loop
        .Close
    End With

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount - 1
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

Rem --- pcr_case_daily.csv ----------------------------------------------------
    Erase OutData
    ReDim OutData(9, 1500)

    InpFileName = "pcr_case_daily.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpDir & "\" & InpFileName
        InpCount = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            InpArray = Split(InpLine, ",")
            If InpCount < 0 Then
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
                InpCount = 1
            Else
                For I = 0 To 9
                    OutData(I, InpCount) = InpArray(I)
                Next
                InpCount = InpCount + 1
            End If
Rem         DoEvents
        Loop
        .Close
    End With

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount - 1
            OutLine = OutData(0, I) & Chr(9) & OutData(1, I) & Chr(9) & OutData(2, I) & Chr(9) & OutData(3, I) & Chr(9) & OutData(4, I) & Chr(9) & OutData(5, I) & Chr(9) & OutData(6, I) & Chr(9) & OutData(7, I) & Chr(9) & OutData(8, I) & Chr(9) & OutData(9, I)
            .WriteText OutLine, 1
        Next
        .SaveToFile OutDir & "\" & OutFileName, 2
        .Close
    End With

Rem ---------------------------------------------------------------------------
    Ret = MsgBox("completed", vbOKOnly)
