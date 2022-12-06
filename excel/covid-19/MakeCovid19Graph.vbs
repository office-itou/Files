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

Rem --- 人口 ------------------------------------------------------------------
    WScript.Echo "開始：初期化データー"

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName = "人口(人口推計2019).csv"
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

    InpFileName = "人口(国勢調査2020).csv"
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

    Rem --- 都道府県ごとの感染者数[日別/7日間平均/10万人あたり] -----------------
    Erase OutValue
    ReDim OutValue(48, UBound(DateList), 4)

    For I = 0 To 4
        OutValue(0, 0, I) = "日付"
        OutValue(1 + 0, 0, I) = "国内合計"
        OutValue(1 + 1, 0, I) = "北海道"
        OutValue(1 + 2, 0, I) = "青森県"
        OutValue(1 + 3, 0, I) = "岩手県"
        OutValue(1 + 4, 0, I) = "宮城県"
        OutValue(1 + 5, 0, I) = "秋田県"
        OutValue(1 + 6, 0, I) = "山形県"
        OutValue(1 + 7, 0, I) = "福島県"
        OutValue(1 + 8, 0, I) = "茨城県"
        OutValue(1 + 9, 0, I) = "栃木県"
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
                    OutValue(I, InpCount + 1, 1) = InpArray(I)                  Rem 7日間平均
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
    For I = OutCount + 1 To UBound(DateList)
        OutDate = DateAdd("d", I - OutCount, DateList(OutCount))
        DateList(   I   ) = OutDate                         Rem 日付一覧
        OutValue(0, I, 0) = OutDate                         Rem 日別
        OutValue(0, I, 1) = OutDate                         Rem 7日間平均
        OutValue(0, I, 2) = OutDate                         Rem 10万人あたり
        OutValue(0, I, 3) = OutDate                         Rem 10万人あたり(7日間平均)
        OutValue(0, I, 4) = OutDate                         Rem 10万人あたり(算出)
    Next
    OutCount = UBound(DateList)
    Rem -----------------------------------------------------------------------
    For I = 0 To 4
        OutFileName = "感染者数." & I & ".txt"
        WScript.Echo "書出：" & OutFileName
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
    Rem --- 順位付け ----------------------------------------------------------
    With CreateObject("ADODB.Recordset")
        .Fields.Append "CD", 200, 128
        .Fields.Append "NAME", 200, 128
        .Fields.Append "VALUE", 5
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
    ReDim OutValue(17, UBound(DateList))

    OutValue(0, 0) = "日付"
    OutValue(1, 0) = "感染者数"
    OutValue(2, 0) = "感染者数(7日間平均)"
    OutValue(3, 0) = "感染者数(10万人あたり)"
    OutValue(4, 0) = "感染者数(10万人あたり・7日間平均)"
    OutValue(5, 0) = "感染者数(累計)"
    OutValue(6, 0) = "死者数"
    OutValue(7, 0) = "死者数(7日間平均)"
    OutValue(8, 0) = "死者数(累計)"
    OutValue(9, 0) = "重症者数"
    OutValue(10, 0) = "重症者数(7日間平均)"
    OutValue(11, 0) = "入院療養中"
    OutValue(12, 0) = "退院療養解除"
    OutValue(13, 0) = "退院療養解除(累計)"
    OutValue(14, 0) = "PCR検査数(自費検査を除く)"
    OutValue(15, 0) = "PCR検査数(自費検査を含む)"
    OutValue(16, 0) = "陽性率"
    OutValue(17, 0) = "陽性率(7日間平均値)"

    For I = 0 To OutCount - 1
        OutValue(0, I + 1) = DateList(I + 1)
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
                For I = 1 To OutCount - 1
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

            OutValue(5, InpCount + 1) = InpArray(1)         Rem 感染者数[累計・全国]
            InpCount = InpCount + 1
        Loop
Rem     OutCount = InpCount
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
                For I = 1 To OutCount - 1
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
            OutValue(8, InpCount + 1) = InpArray(1)         Rem 死者数[累計・全国]
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
                For I = 1 To OutCount - 1
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
                For I = 1 To OutCount - 1
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
                For I = 1 To OutCount - 1
                    If CDate(OutValue(0, I)) = CDate(InpArray(0)) Then
                        InpCount = I - 1
                        Exit For
                    End If
                Next
            End If
            OutValue(14, InpCount + 1) = InpArray(7)        Rem PCR検査数(自費検査を除く)
            OutValue(15, InpCount + 1) = InpArray(9)        Rem PCR検査数(自費検査を含む)
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

    WScript.Echo "終了：データー変換"

Rem --- Excel -----------------------------------------------------------------
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    Set objDstWorkbook = objExcel.Workbooks.Add()

    Rem -----------------------------------------------------------------------
    Call MakeExcelFile("感染者数", "感染者数.0.txt")
    Call MakeExcelFile("7日間平均", "感染者数.1.txt")
    Call MakeExcelFile("10万人", "感染者数.3.txt")
    Call MakeExcelFile("日本国内", "日本国内.txt")
    Call MakeExcelFile("順位付け", "順位付け.txt")

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("Sheet1")
Rem     .Visible = False
        .Activate
        .Name = "グラフ"
        Rem ---  1: 感染者数 --------------------------------------------------
        Rem ---  2: 7日間平均 -------------------------------------------------
        Rem ---  3: 10万人あたりの感染者数 ------------------------------------
        With objDstWorkbook
            Call MakeGraph1(.Worksheets("グラフ"), .Worksheets("感染者数"), .Worksheets("グラフ").ChartObjects.Add(0, 0, 912, 585), "描画：1: 感染者数", "感染者数", False)
            Call MakeGraph1(.Worksheets("グラフ"), .Worksheets("7日間平均"), .Worksheets("グラフ").ChartObjects.Add(960, 0, 912, 585), "描画：2: 7日間平均", "感染者数の7日間平均", False)
            Call MakeGraph1(.Worksheets("グラフ"), .Worksheets("10万人"), .Worksheets("グラフ").ChartObjects.Add(0, 600, 912, 585), "描画：3: 10万人あたりの感染者数", "感染者数の10万人あたり7日間平均", False)
        End With
        Rem ---  4: 感染者数（東京都） ----------------------------------------
        Erase Collection
        ReDim Collection(2, 4)
        Collection(0, 0) = "=""感染者数"""
        Collection(0, 1) = "=感染者数!$A$2:$A$"
        Collection(0, 2) = "=感染者数!$O$2:$O$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""7日間平均"""
        Collection(1, 1) = "=感染者数!$A$2:$A$"
        Collection(1, 2) = "=7日間平均!$O$2:$O$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""10万人"""
        Collection(2, 1) = "=感染者数!$A$2:$A$"
        Collection(2, 2) = "=10万人!$O$2:$O$"
        Collection(2, 3) = 2
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("グラフ"), .Worksheets("感染者数"), .Worksheets("グラフ").ChartObjects.Add(960, 600, 912, 585), "描画：4: 感染者数（東京都）", "感染者数（東京都）", False, Collection)
        End With
        Rem ---  5: 感染者数（日本国内） --------------------------------------
        Erase Collection
        ReDim Collection(2, 4)
        Collection(0, 0) = "=""感染者数"""
        Collection(0, 1) = "=日本国内!$A$2:$A$"
        Collection(0, 2) = "=日本国内!$B$2:$B$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""7日間平均"""
        Collection(1, 1) = "=日本国内!$A$2:$A$"
        Collection(1, 2) = "=日本国内!$C$2:$C$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""10万人"""
        Collection(2, 1) = "=日本国内!$A$2:$A$"
        Collection(2, 2) = "=日本国内!$E$2:$E$"
        Collection(2, 3) = 2
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("グラフ"), .Worksheets("日本国内"), .Worksheets("グラフ").ChartObjects.Add(0, 1200, 912, 585), "描画：5: 感染者数（日本国内）", "感染者数（日本国内）", False, Collection)
        End With
        Rem ---  6: 死者重症者数 ----------------------------------------------
        Erase Collection
        ReDim Collection(3, 4)
        Collection(0, 0) = "=""死者数"""
        Collection(0, 1) = "=日本国内!$A$2:$A$"
        Collection(0, 2) = "=日本国内!$G$2:$G$"
        Collection(0, 3) = 1
        Collection(1, 0) = "=""死者数(7日間平均)"""
        Collection(1, 1) = "=日本国内!$A$2:$A$"
        Collection(1, 2) = "=日本国内!$H$2:$H$"
        Collection(1, 3) = 1
        Collection(2, 0) = "=""重症者数"""
        Collection(2, 1) = "=日本国内!$A$2:$A$"
        Collection(2, 2) = "=日本国内!$J$2:$J$"
        Collection(2, 3) = 1
        Collection(3, 0) = "=""重症者数(7日間平均)"""
        Collection(3, 1) = "=日本国内!$A$2:$A$"
        Collection(3, 2) = "=日本国内!$K$2:$K$"
        Collection(3, 3) = 1
        With objDstWorkbook
            Call MakeGraph2(.Worksheets("グラフ"), .Worksheets("日本国内"), .Worksheets("グラフ").ChartObjects.Add(960, 1200, 912, 585), "描画：6: 死者重症者数", "死者重症者数", False, Collection)
        End With
        Rem ---  7: 直近7日間の人口10万人当たりの新規感染者数 -----------------
        With objDstWorkbook
            MakeGraph1 .Worksheets("グラフ"), .Worksheets("10万人"), .Worksheets("グラフ").ChartObjects.Add(0, 1800, 912, 585), "描画：7: 直近7日間の人口10万人当たりの新規感染者数", "感染者数の10万人あたり7日間平均（抜粋）", True
            Rem --- テキストボックスの描画 ------------------------------------
            objExcel.Application.ScreenUpdating = False
            With .Worksheets("グラフ").ChartObjects(.Worksheets("グラフ").ChartObjects.Count).Chart
                With .Shapes.AddLabel(1, 0, 0, 72, 72)
                    With .TextFrame.Characters
                        .Text = "グラフは厚生労働省のデーター、一覧表は算出のため一致しません"
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
                    .Left = objDstWorkbook.Worksheets("グラフ").ChartObjects(objDstWorkbook.Worksheets("グラフ").ChartObjects.Count).Width - .Width - 12
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
                    .Left = objDstWorkbook.Worksheets("グラフ").ChartObjects(objDstWorkbook.Worksheets("グラフ").ChartObjects.Count).Width - .Width - 12
                End With
            End With
            objExcel.Application.ScreenUpdating = True
        End With
        Rem -------------------------------------------------------------------
Rem     .Visible = True
        .Activate
        Rem --- グラフ保存 ----------------------------------------------------
        I = 0
        For Each objChart In .ChartObjects
            I = I + 1
            OutFileName = PicDir & "\pic" & I & "." & .Shapes("グラフ " & I).AlternativeText & ".png"
            WScript.Echo "保存：" & OutFileName
            .Range(objChart.TopLeftCell.Address(False, False)).Select
            objChart.Chart.Export OutFileName
        Next
    End With

    Rem -----------------------------------------------------------------------
    For Each objWorksheet In objDstWorkbook.Worksheets
        With objWorksheet
            .Activate
            If .Range("A1").Text = "日付" Then
                .Range("B" & (.Cells(.Rows.Count, 2).End(-4162).Row)).Select
            Else
                .Range("A1").Select
            End If
        End With
    Next

    Rem -----------------------------------------------------------------------
    OutFileName = CurDir & "\" & "covid-19.xlsx"
    WScript.Echo "保存：" & OutFileName
    objDstWorkbook.SaveAs (OutFileName)

    Rem -----------------------------------------------------------------------
    CopyExcel2Excel

    Rem --- 順位付けコピペ用 --------------------------------------------------
    With objDstWorkbook.Worksheets("順位付け")
        RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
        .Activate
        .Range("E2:E" & RowsEnd).Copy
    End With

    Rem -----------------------------------------------------------------------
    With objDstWorkbook.Worksheets("グラフ")
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
    WScript.Echo "抽出：" & InpFileName
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
            Case "順位付け"
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
            Case "感染者数"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0"
            Case "7日間平均"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.00"
            Case "10万人"
                .Range("A1:AW1").HorizontalAlignment = -4108
                .Range("A1:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
                .Range("B2:AW" & (.Cells(.Rows.Count, 1).End(-4162).Row)).NumberFormatLocal = "0.0"
            Case "日本国内"
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
            Case "順位付け"
                .Range("A1:E1").HorizontalAlignment = -4108
                .Range("A1:E" & (.Cells(.Rows.Count, 1).End(-4162).Row)).ShrinkToFit = True
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
    Dim I                               'For Next用
    Dim RowsEnd                         'データーの存在する最終行
    Dim objRange1                       'データーソース範囲
    Dim objRange2                       '最大値の検索範囲
    Dim MaxValue                        '最大値の値
    Dim MaxColumn                       '最大値の列
    Dim MaxRow                          '最大値の行
    Dim LatestValue                     '最新の値
    Dim LatestColumn                    '最新の列
    Dim LatestRow                       '最新の行
    Dim RecordsetCount                  'データーラベル件数
    Dim Point(1, 47, 3)                 'データーラベル情報
    Dim objRecordset(1)                 'ソート用オブジェクト
    Dim objWorksheetRank                '
    Dim objRangeRank                    '

    WScript.Echo MessageText

    Rem --- 順位付け関連 ------------------------------------------------------
    Set objWorksheetRank = objDstWorkbook.Worksheets("順位付け")
    Set objRangeRank = objWorksheetRank.Range("B2:B4")
    Rem --- データーソース関連 ------------------------------------------------
    With objSrcWorksheet
        LatestRow = .Cells(.Rows.Count, 2).End(-4162).Row   '合計列
        RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row     '日付列
        Set objRange1 = .Range("A1:AW" & RowsEnd)
        Set objRange2 = .Range("C2:AW" & RowsEnd)
    End With
    Rem --- グラフの描画 ------------------------------------------------------
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
            Rem --- グラフの描画 ----------------------------------------------
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
            Rem --- 選択されたグラフの描画 ------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 1 To .FullSeriesCollection.Count
                If .FullSeriesCollection(I).IsFiltered = False Then
                    Rem --- 選択されたグラフの最大値の位置を取得 --------------
                    With objSrcWorksheet
                        Set objRange2 = .Range(.Cells(2, I + 1), .Cells(RowsEnd, I + 1))
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- データーラベルの描画 (最大値) ---------------------
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
                    Rem --- データーラベルの描画 (最新値) ---------------------
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
            Rem --- データーラベルの取得 --------------------------------------
            objExcel.Application.ScreenUpdating = False
            RecordsetCount = 0
            For I = 1 To .FullSeriesCollection.Count
                If .FullSeriesCollection(I).IsFiltered = False Then
                    Rem --- 選択されたグラフの最大値の位置を取得 --------------
                    With objSrcWorksheet
                        Set objRange2 = .Range(.Cells(2, I + 1), .Cells(RowsEnd, I + 1))
                        MaxValue = objExcel.Max(objRange2)
                        With objRange2.Find(MaxValue, , -4123, 1)
                            MaxColumn = .Column - 1
                            MaxRow = .Row - 1
                        End With
                    End With
                    Rem --- データーラベルの取得 (最大値) ---------------------
                    With .FullSeriesCollection(I).Points(MaxRow)
                        With .DataLabel
                            Point(0, RecordsetCount, 0) = I
                            Point(0, RecordsetCount, 1) = MaxRow
                            Point(0, RecordsetCount, 2) = .Top
                            Point(0, RecordsetCount, 3) = .Height
                        End With
                    End With
                    Rem --- データーラベルの取得 (最新値) ---------------------
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
            Rem --- データーラベルの調整 (初期化) -----------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To 1
                Set objRecordset(I) = CreateObject("ADODB.Recordset")
                With objRecordset(I)
                    .Fields.Append "CD", 5
                    .Fields.Append "POINT", 5
                    .Fields.Append "TOP", 5
                    .Fields.Append "HEIGHT", 5
                    .Open
                    Rem --- データーラベルの調整 (取得) -----------------------
                    For J = 0 To RecordsetCount - 1
                        .AddNew
                        .Fields("CD").Value = Point(I, J, 0)
                        .Fields("POINT").Value = Point(I, J, 1)
                        .Fields("TOP").Value = Point(I, J, 2)
                        .Fields("HEIGHT").Value = Point(I, J, 3)
                        .Sort = "TOP DESC,CD"
                    Next
                    .MoveFirst
                    Rem --- データーラベルの調整 (設定) -----------------------
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
            Rem --- 表示範囲の設定 --------------------------------------------
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
    Dim I                               'For Next用
    Dim RowsEnd                         'データーの存在する最終行
    Dim objRange1                       'データーソース範囲
    Dim objRange2                       '最大値の検索範囲
    Dim MaxValue                        '最大値の値
    Dim MaxColumn                       '最大値の列
    Dim MaxRow                          '最大値の行
    Dim LatestValue                     '最新の値
    Dim LatestColumn                    '最新の列
    Dim LatestRow                       '最新の行
    Dim RecordsetCount                  'データーラベル件数
    Dim Point(1, 47, 3)                 'データーラベル情報
    Dim objRecordset(1)                 'ソート用オブジェクト
    Dim aryStrings                      '

    WScript.Echo MessageText
    Rem --- データーソース関連 ------------------------------------------------
    With objSrcWorksheet
        LatestRow = .Cells(.Rows.Count, 2).End(-4162).Row   '合計列
        RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row     '日付列
        Set objRange1 = Nothing
        Set objRange2 = Nothing
    End With
    Rem --- グラフの描画 ------------------------------------------------------
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
            Rem --- グラフの描画 ----------------------------------------------
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
            Rem --- 選択されたグラフの描画 ------------------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To .FullSeriesCollection.Count - 1
                If .FullSeriesCollection(I + 1).IsFiltered = False Then
                    Rem --- 選択されたグラフの最大値の位置を取得 --------------
                    With .FullSeriesCollection(I + 1)
                        aryStrings = Split(Mid(Collection(I, 2), 2), "!")
                        If aryStrings(0) = "日本国内" And Left(aryStrings(1), 2) = "$G" Then
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
                    Rem --- データーラベルの描画 (最大値) ---------------------
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
                    Rem --- データーラベルの描画 (最新値) ---------------------
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
            Rem --- データーラベルの取得 --------------------------------------
            objExcel.Application.ScreenUpdating = False
            RecordsetCount = 0
            For I = 0 To .FullSeriesCollection.Count - 1
                If .FullSeriesCollection(I + 1).IsFiltered = False Then
                    Rem --- 選択されたグラフの最大値の位置を取得 --------------
                    With .FullSeriesCollection(I + 1)
                        aryStrings = Split(Mid(Collection(I, 2), 2), "!")
                        If aryStrings(0) = "日本国内" And Left(aryStrings(1), 2) = "$G" Then
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
                    Rem --- データーラベルの取得 (最大値) ---------------------
                    With .FullSeriesCollection(I + 1).Points(MaxRow)
                        With .DataLabel
                            Point(0, RecordsetCount, 0) = I + 1
                            Point(0, RecordsetCount, 1) = MaxRow
                            Point(0, RecordsetCount, 2) = .Top
                            Point(0, RecordsetCount, 3) = .Height
                        End With
                    End With
                    Rem --- データーラベルの取得 (最新値) ---------------------
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
            Rem --- データーラベルの調整 (初期化) -----------------------------
            objExcel.Application.ScreenUpdating = False
            For I = 0 To 1
                Set objRecordset(I) = CreateObject("ADODB.Recordset")
                With objRecordset(I)
                    .Fields.Append "CD", 5
                    .Fields.Append "POINT", 5
                    .Fields.Append "TOP", 5
                    .Fields.Append "HEIGHT", 5
                    .Open
                    Rem --- データーラベルの調整 (取得) -----------------------
                    For J = 0 To RecordsetCount - 1
                        .AddNew
                        .Fields("CD").Value = Point(I, J, 0)
                        .Fields("POINT").Value = Point(I, J, 1)
                        .Fields("TOP").Value = Point(I, J, 2)
                        .Fields("HEIGHT").Value = Point(I, J, 3)
                        .Sort = "TOP DESC,CD"
                    Next
                    .MoveFirst
                    Rem --- データーラベルの調整 (設定) -----------------------
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
            Rem --- 表示範囲の設定 --------------------------------------------
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
            Ret = MsgBox("データーのコピーをしますか？", vbYesNo)
            If Ret = 6 Then
                WScript.Echo "転送：Covid19Data.xlsx→covid-19.xlsx"
                Rem --- 感染者数 ----------------------------------------------
                With objDstWorkbook.Worksheets("感染者数")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("感染者数").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("感染者数")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 7日間平均 ---------------------------------------------
                With objDstWorkbook.Worksheets("7日間平均")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("7日間平均").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("7日間平均")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 10万人 ------------------------------------------------
                With objDstWorkbook.Worksheets("10万人")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("10万人").Range("A1:AW" & RowsEnd).Value = .Range("A1:AW" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("10万人")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 日本国内 ----------------------------------------------
                With objDstWorkbook.Worksheets("日本国内")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("日本国内").Range("A1:O" & RowsEnd).Value = .Range("A1:O" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("日本国内")
                    .Activate
                    .Range("A" & RowsEnd).Select
                End With
                Rem --- 順位付け ----------------------------------------------
                With objDstWorkbook.Worksheets("順位付け")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.Worksheets("順位付け").Range("A1:E" & RowsEnd).Value = .Range("A1:E" & RowsEnd).Value
                End With
                With objOrgWorkbook.Worksheets("順位付け")
                    .Activate
                    .Range("A1").Select
                End With
                Rem --- グラフ用 ----------------------------------------------
                With objDstWorkbook.Worksheets("順位付け")
                    RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
                    .Range("E2:E" & RowsEnd).Copy
                End With
                Rem --- 終了処理 ----------------------------------------------
                With objOrgWorkbook.Worksheets("グラフ")
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
