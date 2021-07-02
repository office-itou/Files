Option Explicit

    Dim InpFileName
    Dim InpLine
    Dim InpArray
    Dim InpDate
    Dim InpCode
    Dim InpName
    Dim InpValue
    Dim OldCode
    Dim OutCode(47)
    Dim OutName(47)
    Dim OutDate(47, 1500)
    Dim OutValue(47, 1500)
    Dim OutFileName
    Dim OutLine
    Dim InpCount
    Dim Ret
    Dim I
    Dim J

Rem With Application.FileDialog(msoFileDialogOpen)
Rem     .AllowMultiSelect = False
Rem     .Show
Rem     InpFileName = .SelectedItems(1)
Rem End With

    InpFileName = ".\nhk_news_covid19_prefectures_daily_data.csv"

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile InpFileName
        InpCode = 0
        OldCode = -1
        InpCount = -1
        Do Until .EOS
            InpLine = .ReadText(-2)
            If InpCount >= 0 Then
                InpArray = Split(InpLine, ",")
                InpDate = InpArray(0)
                InpCode = InpArray(1)
                InpName = InpArray(2)
                InpValue = InpArray(3)
                If OldCode <> InpCode Then
                    InpCount = 0
                    OldCode = InpCode
                End If
                OutCode(InpCode) = InpCode
                OutName(InpCode) = InpName
                OutDate(InpCode, InpCount) = InpDate
                OutValue(InpCode, InpCount) = InpValue
            End If
            InpCount = InpCount + 1
Rem         DoEvents
        Loop
        .Close
    End With

    OutFileName = InpFileName & ".txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        For I = 0 To InpCount
            OutLine = ""
            If I = 0 Then
                For J = 1 To 47
                    If OutLine = "" Then
                        OutLine = "“ú•t" & "," & OutName(J)
                    Else
                        OutLine = OutLine & "," & OutName(J)
                    End If
                Next
                .WriteText OutLine, 1
            End If
            OutLine = ""
            For J = 1 To 47
                If OutDate(J, I) <> "" Then
                    If OutLine = "" Then
                        OutLine = OutDate(J, I) & "," & OutValue(J, I)
                    Else
                        OutLine = OutLine & "," & OutValue(J, I)
                    End If
                End If
            Next
            If OutLine <> "" Then
                .WriteText OutLine, 1
            End If
        Next
        .SaveToFile OutFileName, 2
        .Close
    End With

    Ret = MsgBox("completed", vbOKOnly)
