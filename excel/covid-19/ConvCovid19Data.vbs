Rem ---------------------------------------------------------------------------
Rem ConvCovid19Data.vbs: CSV->TXT(TAB区切り)変換
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

Rem --- 人口 ------------------------------------------------------------------
    Wscript.Echo "開始：初期化データー"

    Erase Population
    ReDim Population(3, 48, 2)

    InpFileName="人口(人口推計2019).csv"
    Wscript.Echo "読出：" & InpFileName
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

    InpFileName="人口(国勢調査2020).csv"
    Wscript.Echo "読出：" & InpFileName
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

    Wscript.Echo "終了：初期化データー"

Rem --- severe_cases_daily ----------------------------------------------------
Rem 重症者数の推移
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

    Wscript.Echo "書出：" & OutFileName
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
Rem 入院治療等を要する者等推移
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

    Wscript.Echo "書出：" & OutFileName
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
Rem PCR検査実施人数
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

    Wscript.Echo "書出：" & OutFileName
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

Rem --- データー取得 ----------------------------------------------------------
    Erase OutValue
    ReDim OutValue(49, 1999, 4)
    Rem 0: 各地の感染者数_1日ごとの発表数
    Rem 1: 各地の死者数_1日ごとの発表数
    Rem 2: 各地の直近1週間の人口10万人あたりの感染者数
    Rem 3: 各地の直近1週間の感染者数(算出)
    Rem 4: 各地の直近1週間の人口10万人あたりの感染者数(算出)
    Rem タイトル行
    For I = 0 To 4
        OutValue(0, 0, I) = "日付"
        OutValue(1, 0, I) = "国内合計"
        OutValue(2, 0, I) = "空港検疫など"
        OutValue(2 +  1, 0, I) = "01:北海道"
        OutValue(2 +  2, 0, I) = "02:青森県"
        OutValue(2 +  3, 0, I) = "03:岩手県"
        OutValue(2 +  4, 0, I) = "04:宮城県"
        OutValue(2 +  5, 0, I) = "05:秋田県"
        OutValue(2 +  6, 0, I) = "06:山形県"
        OutValue(2 +  7, 0, I) = "07:福島県"
        OutValue(2 +  8, 0, I) = "08:茨城県"
        OutValue(2 +  9, 0, I) = "09:栃木県"
        OutValue(2 + 10, 0, I) = "10:群馬県"
        OutValue(2 + 11, 0, I) = "11:埼玉県"
        OutValue(2 + 12, 0, I) = "12:千葉県"
        OutValue(2 + 13, 0, I) = "13:東京都"
        OutValue(2 + 14, 0, I) = "14:神奈川県"
        OutValue(2 + 15, 0, I) = "15:新潟県"
        OutValue(2 + 16, 0, I) = "16:富山県"
        OutValue(2 + 17, 0, I) = "17:石川県"
        OutValue(2 + 18, 0, I) = "18:福井県"
        OutValue(2 + 19, 0, I) = "19:山梨県"
        OutValue(2 + 20, 0, I) = "20:長野県"
        OutValue(2 + 21, 0, I) = "21:岐阜県"
        OutValue(2 + 22, 0, I) = "22:静岡県"
        OutValue(2 + 23, 0, I) = "23:愛知県"
        OutValue(2 + 24, 0, I) = "24:三重県"
        OutValue(2 + 25, 0, I) = "25:滋賀県"
        OutValue(2 + 26, 0, I) = "26:京都府"
        OutValue(2 + 27, 0, I) = "27:大阪府"
        OutValue(2 + 28, 0, I) = "28:兵庫県"
        OutValue(2 + 29, 0, I) = "29:奈良県"
        OutValue(2 + 30, 0, I) = "30:和歌山県"
        OutValue(2 + 31, 0, I) = "31:鳥取県"
        OutValue(2 + 32, 0, I) = "32:島根県"
        OutValue(2 + 33, 0, I) = "33:岡山県"
        OutValue(2 + 34, 0, I) = "34:広島県"
        OutValue(2 + 35, 0, I) = "35:山口県"
        OutValue(2 + 36, 0, I) = "36:徳島県"
        OutValue(2 + 37, 0, I) = "37:香川県"
        OutValue(2 + 38, 0, I) = "38:愛媛県"
        OutValue(2 + 39, 0, I) = "39:高知県"
        OutValue(2 + 40, 0, I) = "40:福岡県"
        OutValue(2 + 41, 0, I) = "41:佐賀県"
        OutValue(2 + 42, 0, I) = "42:長崎県"
        OutValue(2 + 43, 0, I) = "43:熊本県"
        OutValue(2 + 44, 0, I) = "44:大分県"
        OutValue(2 + 45, 0, I) = "45:宮崎県"
        OutValue(2 + 46, 0, I) = "46:鹿児島県"
        OutValue(2 + 47, 0, I) = "47:沖縄県"
    Next

Rem --- deaths_cumulative_daily -----------------------------------------------
Rem 死亡者数（累積）
    InpFileName="deaths_cumulative_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

    Erase OutData
    ReDim OutData(49)

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
            For I = 0 To 48
                Select Case I
                    Case 0              Rem 日付
                        OutValue(0, InpCount + 1, 1) = InpArray(I)
                    Case 1              Rem 国内合計
                        If InpCount > 0 Then
                            OutValue(1, InpCount + 1, 1) = InpArray(I) - OutData(I)
                        Else
                            OutValue(1, InpCount + 1, 1) = InpArray(I)
                        End If
                        OutData(I) = InpArray(I)
                    Case Else           Rem 各地の1日ごとの発表数
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

    Rem --- 各地の死者数_1日ごとの発表数 --------------------------------------
    Wscript.Echo "書出：" & OutFileName

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
Rem 新規陽性者数の推移（日別）
    InpFileName="newly_confirmed_cases_daily.csv"
    OutFileName = InpFileName & ".txt"
    Wscript.Echo "開始：" & InpFileName

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
            For I = 0 To 48
                Select Case I
                    Case 0          Rem 日付
                        OutValue(0, InpCount + 1, 0) = InpArray(I)
                        OutValue(0, InpCount + 1, 3) = InpArray(I)
                        OutValue(0, InpCount + 1, 4) = InpArray(I)
                    Case 1          Rem 国内合計
                        OutValue(1, InpCount + 1, 0) = InpArray(I)
                    Case Else       Rem 各地の1日ごとの発表数
                        OutValue(1 + I, InpCount + 1, 0) = InpArray(I)
                End Select
            Next
            InpCount = InpCount + 1
        Loop
        .Close
    End With

    Rem --- 各地の直近1週間の感染者数の算出 -----------------------------------
    Wscript.Echo "計算：" & InpFileName
    For I = 0 To InpCount
        If I >= 6 Then
            For J = 1 To 49
                InpValue = 0
                Rem 7日間合計
                For K = 0 To 6
                    InpValue = InpValue + OutValue(J, I + 1 - K, 0)
                Next
                Rem 7日間平均
                OutValue(J, I + 1, 3) = Round(InpValue / 7, 2)
                Rem 人口10万人あたり
                Select Case J
                    Case 1          Rem 国内合計
                        If CDate(OutValue(0, I + 1, 0)) < CDate("2022/1/1") Then
                            OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J, 0) * 100000)
                        Else
                            OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J, 1) * 100000)
                        End If
                    Case 2          Rem 空港検疫など
                    Case Else       Rem 各地
                        If CDate(OutValue(0, I + 1, 0)) < CDate("2022/1/1") Then
                            OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J - 1, 0) * 100000)
                        Else
                            OutValue(J, I + 1, 4) = CDbl(InpValue / Population(3, J - 1, 1) * 100000)
                        End If
                End Select
            Next
        End If
    Next
    Rem --- 最新の各地の直近1週間の人口10万人あたりの感染者数の順位付け -------
    With CreateObject("ADODB.Recordset")
        .Fields.Append "CD",200,128
        .Fields.Append "NAME",200,128
        .Fields.Append "VALUE",5
        .Open
        For I = 0 To 46
            .AddNew
            .Fields("CD").Value = Left(OutValue(I + 3, 0, 4), 2)                Rem 都道府県コード
            .Fields("NAME").Value = Mid(OutValue(I + 3, 0, 4), 4)               Rem 都道府県名
            .Fields("VALUE").Value = OutValue(I + 3, InpCount + 0, 4)           Rem 各地の直近1週間の人口10万人あたりの感染者数
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
            RankData(3, I) = FormatNumber(Round(.Fields("VALUE").Value, 2), 2, -1)
            RankData(4, I) = RankData(2, I) & ":" & RankData(3, I)
            .MoveNext
        Next
        .Close
    End With

    Rem --- 0: 各地の感染者数_1日ごとの発表数 ---------------------------------
    Wscript.Echo "書出：" & OutFileName
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
    Rem --- 3: 各地の直近1週間の感染者数 --------------------------------------
    OutFileName = InpFileName & ".3.txt"
    Wscript.Echo "書出：" & OutFileName
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
    Rem --- 4: 各地の直近1週間の人口10万人あたりの感染者数(算出) --------------
    OutFileName = InpFileName & ".4.txt"
    Wscript.Echo "書出：" & OutFileName
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
    Rem --- 最新の各地の直近1週間の人口10万人あたりの感染者数の順位付け -------
    OutFileName = InpFileName & ".4.順位付け.txt"
    Wscript.Echo "書出：" & OutFileName
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
            Case "newly_confirmed_cases_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地感染者"
            Case "deaths_cumulative_daily.csv.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地死者数"
            Case "newly_confirmed_cases_daily.csv.3.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地 7日間"
            Case "newly_confirmed_cases_daily.csv.4.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地10万人(算出)"
            Case "newly_confirmed_cases_daily.csv.4.順位付け.txt"
                objSrcWorkbook.WorkSheets(1).Name = "各地10万人(順位)"
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

    Wscript.Echo "保存：" & CurDir & "\Covid19Data.xlsx"
    objDstWorkbook.SaveAs(CurDir & "\Covid19Data.xlsx")

Rem ---------------------------------------------------------------------------
Rem グラフ      ：各地10万人(順位)
Rem 感染者数    ：各地感染者
Rem 7日間平均値 ：各地 7日間
Rem 10万人あたり：各地10万人/各地10万人(算出)
Rem 日本国内    ：国内感染者/国内重症者/国内入退院/PCR 検査数
    Set obOrgjExcel = GetObject(, "Excel.Application")
    For Each objOrgWorkbook In obOrgjExcel.Workbooks
        If objOrgWorkbook.Name = "covid-19.xlsx" Then
            Ret = MsgBox("データーのコピーをしますか？", vbYesNo)
            If Ret = 6 Then
                Wscript.Echo "転送：Covid19Data.xlsx→covid-19.xlsx"
                Rem --- 感染者数 ----------------------------------------------
                With objDstWorkbook.WorkSheets("各地感染者")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
                    objOrgWorkbook.WorkSheets("感染者数").Range("B" & (RowsEnd + 3 - 2) & ":AX" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":AX" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("感染者数")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- 7日間平均値 -------------------------------------------
                With objDstWorkbook.WorkSheets("各地 7日間")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
Rem                 objOrgWorkbook.WorkSheets("7日間平均値").Range("B" & (RowsEnd + 3 - 2) & ":AX" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":AX" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("7日間平均値")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- 10万人あたり ------------------------------------------
                With objDstWorkbook.WorkSheets("各地10万人(算出)")
                    RowsEnd = .Cells(.Rows.Count, 1).End(-4162).Row
Rem                 objOrgWorkbook.WorkSheets("10万人あたり").Range("B" & (RowsEnd + 3 - 2) & ":B" & (RowsEnd + 3 - 2)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("10万人あたり")
                    .Activate 
                    .Range("B" & (RowsEnd + 3 - 2)).Select
                End With
                Rem --- 日本国内 ----------------------------------------------
                With objDstWorkbook.WorkSheets("国内重症者")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("G" & (RowsEnd + 989 - 874) & ":G" & (RowsEnd + 989 - 874)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("国内入退院")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("I" & (RowsEnd + 117 - 2) & ":I" & (RowsEnd + 117 - 2)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                    RowsEnd = .Cells(.Rows.Count, 3).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("J" & (RowsEnd + 118 - 3) & ":J" & (RowsEnd + 118 - 3)).Value = .Range("C" & RowsEnd & ":C" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("PCR 検査数")
                    RowsEnd = .Cells(.Rows.Count, 10).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("K" & (RowsEnd + 36 - 2) & ":K" & (RowsEnd + 36 - 2)).Value = .Range("J" & RowsEnd & ":J" & RowsEnd).Value
                End With
                With objDstWorkbook.WorkSheets("各地死者数")
                    RowsEnd = .Cells(.Rows.Count, 2).End(-4162).Row
                    objOrgWorkbook.WorkSheets("日本国内").Range("E" & (RowsEnd + 989 - 874) & ":E" & (RowsEnd + 989 - 874)).Value = .Range("B" & RowsEnd & ":B" & RowsEnd).Value
                End With
                With objOrgWorkbook.WorkSheets("日本国内")
                    .Activate 
                    .Range("B" & (RowsEnd + 989 - 874)).Select
                End With
                Rem --- グラフ用 ----------------------------------------------
Rem             With objDstWorkbook.WorkSheets("各地10万人(順位)")
Rem                 RowsEnd = .Cells(.Rows.Count, 5).End(-4162).Row
Rem                 .Range("E2:E" & RowsEnd).Copy
Rem             End With
                Rem --- 終了処理 ----------------------------------------------
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
