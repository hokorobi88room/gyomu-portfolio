Attribute VB_Name = "modValidate"
'==============================================================
' modValidate ― 実行前の入力検査
' 「動く前に壊れているデータを見つける」ための事前検査。
' 結果は「検査結果」シートへ行単位で書き出す(黙って落ちない)。
'==============================================================
Option Explicit

Public Function ValidateAll(ByRef cfg As TConfig) As Boolean
    Dim problems As New Collection

    CheckCustomers problems
    CheckSales problems, cfg

    ' --- 結果シートへ出力 ---
    Dim ws As Worksheet
    Set ws = modLog.EnsureSheet(modConfig.SHEET_CHECK, Array("シート", "行", "内容"))
    ws.Range("A2:C" & ws.Rows.Count).ClearContents

    Dim i As Long, p As Variant
    i = 2
    For Each p In problems
        ws.Cells(i, 1).Value = p(0)
        ws.Cells(i, 2).Value = p(1)
        ws.Cells(i, 3).Value = p(2)
        i = i + 1
    Next p

    If problems.Count = 0 Then
        ws.Cells(2, 1).Value = "-"
        ws.Cells(2, 3).Value = "問題なし(" & Format$(Now, "yyyy/mm/dd hh:nn") & " 検査)"
    End If

    ValidateAll = (problems.Count = 0)
End Function

Private Sub CheckCustomers(ByVal problems As Collection)
    Dim ws As Worksheet, arr As Variant, i As Long
    Set ws = modConfig.SheetByName(modConfig.SHEET_CUSTOMER)
    arr = modData.UsedBody(ws)
    If IsEmpty(arr) Then
        problems.Add Array(modConfig.SHEET_CUSTOMER, 0, "顧客マスタが空です")
        Exit Sub
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(arr, 1)
        Dim rowNo As Long, code As String
        rowNo = i + 1
        code = Trim$(CStr(arr(i, 1)))
        If Len(code) = 0 Then
            If Len(Trim$(CStr(arr(i, 2)))) > 0 Then
                problems.Add Array(modConfig.SHEET_CUSTOMER, rowNo, "顧客コードが空です")
            End If
        ElseIf seen.Exists(code) Then
            problems.Add Array(modConfig.SHEET_CUSTOMER, rowNo, "顧客コード重複: " & code)
        Else
            seen.Add code, True
            If Len(Trim$(CStr(arr(i, 2)))) = 0 Then
                problems.Add Array(modConfig.SHEET_CUSTOMER, rowNo, "顧客名が空です: " & code)
            End If
        End If
    Next i
End Sub

Private Sub CheckSales(ByVal problems As Collection, ByRef cfg As TConfig)
    Dim ws As Worksheet, arr As Variant, i As Long
    Set ws = modConfig.SheetByName(modConfig.SHEET_SALES)
    arr = modData.UsedBody(ws)
    If IsEmpty(arr) Then
        problems.Add Array(modConfig.SHEET_SALES, 0, "売上明細が空です")
        Exit Sub
    End If

    For i = 1 To UBound(arr, 1)
        Dim rowNo As Long
        rowNo = i + 1
        If IsEmpty(arr(i, 1)) And IsEmpty(arr(i, 2)) Then
            ' 完全な空行は許容(スキップ)
        Else
            If Not IsDate(arr(i, 1)) Then _
                problems.Add Array(modConfig.SHEET_SALES, rowNo, "日付が不正: " & CStr(arr(i, 1)))
            If Len(Trim$(CStr(arr(i, 2)))) = 0 Then _
                problems.Add Array(modConfig.SHEET_SALES, rowNo, "顧客コードが空です")
            If Not IsNumeric(arr(i, 4)) Or Val(arr(i, 4)) <= 0 Then _
                problems.Add Array(modConfig.SHEET_SALES, rowNo, "数量が不正: " & CStr(arr(i, 4)))
            If Not IsNumeric(arr(i, 5)) Then _
                problems.Add Array(modConfig.SHEET_SALES, rowNo, "単価が不正: " & CStr(arr(i, 5)))
            Dim rate As String
            rate = Trim$(CStr(arr(i, 6)))
            If rate <> "10" And rate <> "8" Then _
                problems.Add Array(modConfig.SHEET_SALES, rowNo, "税率は 10 か 8 で入力してください: " & rate)
        End If
    Next i
End Sub
