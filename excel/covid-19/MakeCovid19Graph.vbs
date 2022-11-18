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

Rem --- 人口 ------------------------------------------------------------------
    WScript.Echo "開始：初期化データー"

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName="人口(人口推計2019).csv"
    WScript.Echo "読出：" & InpFileName

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

    InpFileName="人口(国勢調査2020).csv"
    WScript.Echo "読出：" & InpFileName

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

    WScript.Echo "終了：初期化データー"

Rem --- データー変換 ----------------------------------------------------------
    WScript.Echo "開始：データー変換"

    Erase DateList
    ReDim DateList(1999)

    DateList(0) = "日付"

    Rem --- 都道府県ごとの感染者数[日別/7日平均/10万人あたり] -----------------
    Erase OutValue
    ReDim OutValue(48, 1999, 4)

    For I = 0 To 4
        OutValue(0     , 0, I) = "日付"
        OutValue(1 +  0, 0, I) = "国内合計"
        OutValue(1 +  1, 0, I) = "北海道"
        OutValue(1 +  2, 0, I) = "青森県"
        OutValue(1 +  3, 0, I) = "岩手県"
        OutValue(1 +  4, 0, I) = "宮城県"
        OutValue(1 +  5, 0, I) = "秋田県"
        OutValue(1 +  6, 0, I) = "山形県"
        OutValue(1 +  7, 0, I) = "福島県"
        OutValue(1 +  8, 0, I) = "茨城県"
        OutValue(1 +  9, 0, I) = "栃木県"
        OutValue(1 + 10, 0, I) = "群馬県"
        OutValue(1 + 11, 0, I) = "埼玉県"
        OutValue(1 + 12, 0, I) = "千葉県"
        OutValue(1 + 13, 0, I) = "東京都"
        OutValue(1 + 14, 0, I) = "神奈川県"
        OutValue(1 + 15, 0, I) = "新潟県"
        OutValue(1 + 16, 0, I) = "富山県"
        OutValue(1 + 17, 0, I) = "石川県"
        OutValue(1 + 18, 0, I) = "福井県"
        OutValue(1 + 19, 0, I) = "山梨県"
        OutValue(1 + 20, 0, I) = "長野県"
        OutValue(1 + 21, 0, I) = "岐阜県"
        OutValue(1 + 22, 0, I) = "静岡県"
        OutValue(1 + 23, 0, I) = "愛知県"
        OutValue(1 + 24, 0, I) = "三重県"
        OutValue(1 + 25, 0, I) = "滋賀県"
        OutValue(1 + 26, 0, I) = "京都府"
        OutValue(1 + 27, 0, I) = "大阪府"
        OutValue(1 + 28, 0, I) = "兵庫県"
        OutValue(1 + 29, 0, I) = "奈良県"
        OutValue(1 + 30, 0, I) = "和歌山県"
        OutValue(1 + 31, 0, I) = "鳥取県"
        OutValue(1 + 32, 0, I) = "島根県"
        OutValue(1 + 33, 0, I) = "岡山県"
        OutValue(1 + 34, 0, I) = "広島県"
        OutValue(1 + 35, 0, I) = "山口県"
        OutValue(1 + 36, 0, I) = "徳島県"
        OutValue(1 + 37, 0, I) = "香川県"
        OutValue(1 + 38, 0, I) = "愛媛県"
        OutValue(1 + 39, 0, I) = "高知県"
        OutValue(1 + 40, 0, I) = "福岡県"
        OutValue(1 + 41, 0, I) = "佐賀県"
        OutValue(1 + 42, 0, I) = "長崎県"
        OutValue(1 + 43, 0, I) = "熊本県"
        OutValue(1 + 44, 0, I) = "大分県"
        OutValue(1 + 45, 0, I) = "宮崎県"
        OutValue(1 + 46, 0, I) = "鹿児島県"
        OutValue(1 + 47, 0, I) = "沖縄県"
    Next
    Rem -----------------------------------------------------------------------
    InpFileName = "newly_confirmed_cases_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
                    DateList(   InpCount + 1   ) = InpArray(I)                  Rem 日付一覧
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  Rem 日別
                    OutValue(I, InpCount + 1, 1) = InpArray(I)                  Rem 7日平均
                    OutValue(I, InpCount + 1, 2) = InpArray(I)                  Rem 10万人あたり
                    OutValue(I, InpCount + 1, 3) = InpArray(I)                  Rem 10万人あたり(7日間平均)
                    OutValue(I, InpCount + 1, 4) = InpArray(I)                  Rem 10万人あたり(算出)
                Else
                    OutValue(I, InpCount + 1, 0) = InpArray(I)                  Rem 日別
                    If InpCount >= 6 Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 0)
                        Next
                        OutValue(I, InpCount + 1, 1) = Round(InpValue / 7, 2)   Rem 7日間平均
                                                                                Rem 10万人あたり(算出)
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
    WScript.Echo "読出：" & InpFileName
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
                OutValue(I, InpCount + 1, 2) = InpArray(I)                      Rem 10万人あたり(日別)
                If InpCount >= 6 Then
                    If IsNumeric(OutValue(I, InpCount + 1 - 7, 2)) = True Then
                        InpValue = 0
                        For J = 0 To 6
                            InpValue = InpValue + OutValue(I, InpCount + 1 - J, 2)
                        Next
                        OutValue(I, InpCount + 1, 3) = InpValue                 Rem 10万人あたり(7日間合計)
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
        OutFileName = "感染者数." & I & ".txt"
        WScript.Echo "書出：" & OutFileName
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
    Rem --- 順位付け ----------------------------------------------------------
    With CreateObject("ADODB.Recordset")
        .Fields.Append "CD",200,128
        .Fields.Append "NAME",200,128
        .Fields.Append "VALUE",5
        .Open
        For I = 0 To 46
            .AddNew
            .Fields("CD").Value = I + 1                                         Rem 都道府県コード
            .Fields("NAME").Value = OutValue(I + 2, 0, 4)                       Rem 都道府県名
            .Fields("VALUE").Value = OutValue(I + 2, InpCount + 0, 4)           Rem 各地の直近1週間の人口10万人あたりの感染者数
        Next
        .Sort = "VALUE DESC,CD"
        .MoveFirst
        Erase RankData
        ReDim RankData(4, 47)
        RankData(0, 0) = "順位"
        RankData(1, 0) = "都道府県CD"
        RankData(2, 0) = "都道府県名"
        RankData(3, 0) = "感染者数"
        RankData(4, 0) = "コピペ用"
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
    OutFileName = "順位付け.txt"
    WScript.Echo "書出：" & OutFileName
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

    Rem --- 日本国内[感染者数/死者数/重症者数/入院療養中/退院療養解除/PCR検査数]
    Erase OutValue
    ReDim OutValue(16, 1999)

    OutValue( 0, 0) = "日付"
    OutValue( 1, 0) = "感染者数"
    OutValue( 2, 0) = "感染者数(7日間平均)"
    OutValue( 3, 0) = "感染者数(10万人あたり)"
    OutValue( 4, 0) = "感染者数(10万人あたり・7日間平均)"
    OutValue( 5, 0) = "感染者数(累計)"
    OutValue( 6, 0) = "死者数"
    OutValue( 7, 0) = "死者数(7日間平均)"
    OutValue( 8, 0) = "死者数(累計)"
    OutValue( 9, 0) = "重症者数"
    OutValue(10, 0) = "重症者数(7日間平均)"
    OutValue(11, 0) = "入院療養中"
    OutValue(12, 0) = "退院療養解除"
    OutValue(13, 0) = "退院療養解除(累計)"
    OutValue(14, 0) = "PCR検査数"
    OutValue(15, 0) = "陽性率"
    OutValue(16, 0) = "陽性率(7日間平均値)"

    For I = 0 To OutCount
        OutValue( 0, I + 1) = DateList(I + 1)
    Next

    Rem --- 感染者数[累計] ----------------------------------------------------
    InpFileName = "confirmed_cases_cumulative_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
                                                           Rem 感染者数[差分・全国]
            If IsNumeric(OutValue(4, InpCount + 1 - 1)) = False Then
                OutValue(1, InpCount + 1) = InpArray(1)
            Else
                OutValue(1, InpCount + 1) = InpArray(1) - OutValue(5, InpCount + 1 - 1)
            End If
                                                                                Rem 10万人あたり(日別)
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
                OutValue(2, InpCount + 1) = Round(InpValue / 7, 2)              Rem 7日間平均
                If IsNumeric(OutValue(3, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For J = 0 To 6
                        InpValue = InpValue + OutValue(3, InpCount + 1 - J)
                    Next
                    OutValue(4, InpCount + 1) = InpValue                        Rem 10万人あたり(7日間合計)
                End If
            End If

            OutValue(5, InpCount + 1) = InpArray(1)        Rem 感染者数[累計・全国]
            InpCount = InpCount + 1
        Loop
        OutCount = InpCount
        .Close
    End With
    Rem --- 死者数 ------------------------------------------------------------
    InpFileName = "deaths_cumulative_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
                                                            Rem 死者数[日別・全国]
            If IsNumeric(OutValue(6, InpCount + 1 - 1)) = False Then
                OutValue(6, InpCount + 1) = InpArray(1)
            Else
                OutValue(6, InpCount + 1) = InpArray(1) - OutData
            End If
            OutData = InpArray(1)
                                                            Rem 7日間平均
            If InpCount >= 6 Then
                If IsNumeric(OutValue(6, InpCount + 1 - 7)) = True Then
                    InpValue = 0
                    For I = 0 To 6
                        InpValue = InpValue + OutValue(6, InpCount + 1 - I)
                    Next
                    OutValue(7, InpCount + 1) = Round(InpValue / 7, 2)
                End If
            End If
            OutValue(8, InpCount + 1) = InpArray(1)        Rem 死者数[累計・全国]
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- 重症者数 ----------------------------------------------------------
    InpFileName = "severe_cases_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
            OutValue(9, InpCount + 1) = InpArray(1)         Rem 重症者数[日別・全国]
                                                            Rem 7日間平均
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
    Rem --- 入院療養中/退院療養解除 -------------------------------------------
    InpFileName = "requiring_inpatient_care_etc_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
            OutValue(11, InpCount + 1) = InpArray(1)        Rem 入院療養中[日別・全国]
                                                            Rem 退院療養解除[日別・全国]
            If IsNumeric(OutValue(12, InpCount + 1 - 1)) = False Then
                OutValue(12, InpCount + 1) = InpArray(2)
            Else
                OutValue(12, InpCount + 1) = InpArray(2) - OutData
            End If
            OutValue(13, InpCount + 1) = InpArray(2)        Rem 退院療養解除[累計・全国]
            OutData = InpArray(2)
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- PCR検査数 ---------------------------------------------------------
    InpFileName = "pcr_case_daily.csv"
    WScript.Echo "読出：" & InpFileName
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
            OutValue(14, InpCount + 1) = InpArray(9)        Rem PCR検査数
            InpCount = InpCount + 1
        Loop
        .Close
    End With
    Rem --- 日本国内 ----------------------------------------------------------
    OutFileName = "日本国内.txt"
    WScript.Echo "書出：" & OutFileName
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

    WScript.Echo "終了：データー変換"

Rem --- Excel -----------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

    Rem -----------------------------------------------------------------------
    MakeExcelFile "感染者数", "感染者数.0.txt"
    MakeExcelFile "7日平均" , "感染者数.1.txt"
    MakeExcelFile "10万人"  , "感染者数.3.txt"
    MakeExcelFile "日本国内", "日本国内.txt"
    MakeExcelFile "順位付け", "順位付け.txt"

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("Sheet1")
        .Activate
        .Name = "グラフ"
    End With

    Rem -----------------------------------------------------------------------
    WScript.Echo "保存：" & CurDir & "\" & "MakeCovid19Graph.xlsx"
    objDstWorkbook.SaveAs(CurDir & "\" & "MakeCovid19Graph.xlsx")

    Rem -----------------------------------------------------------------------
    Set obOrgjExcel = GetObject(, "Excel.Application")
    For Each objOrgWorkbook In obOrgjExcel.Workbooks
        If objOrgWorkbook.Name = "covid-19.xlsx" Then
            Ret = MsgBox("データーのコピーをしますか？", vbYesNo)
            If Ret = 6 Then
                Wscript.Echo "転送：Covid19Data.xlsx→covid-19.xlsx"
                Rem --- 感染者数 ----------------------------------------------
                With objDstWorkbook.WorkSheets("感染者数")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("感染者数").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("感染者数")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 7日平均 -----------------------------------------------
                With objDstWorkbook.WorkSheets("7日平均")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("7日平均").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("7日平均")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 10万人 ------------------------------------------------
                With objDstWorkbook.WorkSheets("10万人")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("10万人").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("10万人")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 日本国内 ----------------------------------------------
                With objDstWorkbook.WorkSheets("日本国内")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("A1:O" & RowsEnd).Value = .Range("A1:O" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("日本国内")
                    .Activate 
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 順位付け ----------------------------------------------
                With objDstWorkbook.WorkSheets("順位付け")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("順位付け").Range("A1:E" & RowsEnd).Value = .Range("A1:E" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("順位付け")
                    .Activate 
                    .Range("A1").Select
                End With
                Rem --- グラフ用 ----------------------------------------------
                With objDstWorkbook.WorkSheets("順位付け")
                    RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
                    .Range("E2:E" & RowsEnd).Copy
                End With
                Rem --- 終了処理 ----------------------------------------------
                With objOrgWorkbook.WorkSheets("グラフ")
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
    WScript.Echo "抽出：" & InpFileName
    objExcel.Workbooks.OpenText OutDir & "\" & InpFileName,65001,,,,,True
    Set objSrcWorkbook = objExcel.Workbooks.Item(objExcel.Workbooks.Count)
    objSrcWorkbook.WorkSheets(1).Name = WorkSheetName
    With objExcel
        objSrcWorkbook.WorkSheets(1).Select
        .ActiveWindow.FreezePanes = False
        .Range("B2").Select
        .ActiveWindow.FreezePanes = True
        Select Case WorkSheetName
            Case "感染者数"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
            Case "7日平均"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "10万人"
                With .Range("A1:AW1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "日本国内"
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
            Case "順位付け"
                With .Range("A1:E1")
                    .HorizontalAlignment = -4108
                    .ShrinkToFit = True
                End With
                .Range("D2:D" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
        End Select
        If .Range("A1").Text = "日付" Then
            With .Range("A2:A" & (.Cells(.Rows.Count, 1).End(-4162).Row))
                .NumberFormatLocal = "yyyy/mm/dd(aaa)"
                .HorizontalAlignment = -4108
            End With
            .Columns("A").AutoFit
        End If
        .Cells.EntireRow.AutoFit
        If .Range("A1").Text = "日付" Then
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
