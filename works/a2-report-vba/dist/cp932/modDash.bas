Attribute VB_Name = "modDash"
'==============================================================
' modDash ― ダッシュボードシートの自動構築
' 月次推移(前年同月比つき)・担当者別ランキング・商品別TOP5 を
' 数式ではなくVBAで集計して書き込む(報告ブック側の構造に依存しない)。
'==============================================================
Option Explicit

Public Sub BuildDashboard(ByVal rows As Collection)
    Dim ws As Worksheet
    Set ws = modUtil.EnsureSheet("ダッシュボード", Array())
    ws.Cells.Clear

    ' --- 集計辞書 ---
    Dim byMonth As Object, byStaff As Object, byItem As Object
    Set byMonth = CreateObject("Scripting.Dictionary")
    Set byStaff = CreateObject("Scripting.Dictionary")
    Set byItem = CreateObject("Scripting.Dictionary")

    Dim item As Variant
    For Each item In rows
        Dim ym As String
        ym = Format$(item(0), "yyyy/mm")
        byMonth(ym) = modUtil.DictNum(byMonth, ym) + item(3)
        byStaff(item(1)) = modUtil.DictNum(byStaff, CStr(item(1))) + item(3)
        byItem(item(2)) = modUtil.DictNum(byItem, CStr(item(2))) + item(3)
    Next item

    ' --- タイトル ---
    ws.Range("A1").Value = "売上ダッシュボード(自動生成: " & Format$(Now, "yyyy/mm/dd hh:nn") & ")"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ' --- 月次推移(前年同月比) ---
    ws.Range("A3").Value = "■ 月次推移"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4:D4").Value = Array("月", "売上", "前年同月", "前年比")
    ws.Range("A4:D4").Font.Bold = True

    Dim keys As Variant, i As Long, r As Long
    keys = SortedKeys(byMonth)
    r = 5
    For i = LBound(keys) To UBound(keys)
        Dim ym2 As String, prevYm As String
        ym2 = keys(i)
        prevYm = Format$(DateAdd("yyyy", -1, CDate(ym2 & "/01")), "yyyy/mm")
        ws.Cells(r, 1).Value = ym2
        ws.Cells(r, 2).Value = byMonth(ym2)
        If byMonth.Exists(prevYm) Then
            ws.Cells(r, 3).Value = byMonth(prevYm)
            ws.Cells(r, 4).Value = byMonth(ym2) / byMonth(prevYm) - 1
        End If
        r = r + 1
    Next i
    ws.Range(ws.Cells(5, 2), ws.Cells(r - 1, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(5, 4), ws.Cells(r - 1, 4)).NumberFormat = "+0.0%;-0.0%"
    AddDataBars ws.Range(ws.Cells(5, 2), ws.Cells(r - 1, 2))
    MarkNegatives ws.Range(ws.Cells(5, 4), ws.Cells(r - 1, 4))

    ' --- 担当者ランキング ---
    Dim startCol As Long
    startCol = 6
    ws.Cells(3, startCol).Value = "■ 担当者別ランキング"
    ws.Cells(3, startCol).Font.Bold = True
    WriteRanking ws, byStaff, 4, startCol, 0   ' 全員
    ' --- 商品TOP5 ---
    ws.Cells(3, startCol + 3).Value = "■ 商品別 TOP5"
    ws.Cells(3, startCol + 3).Font.Bold = True
    WriteRanking ws, byItem, 4, startCol + 3, 5

    ws.Columns("A:L").AutoFit
End Sub

' 辞書を値の降順で書き出す(topN=0 なら全件)
Private Sub WriteRanking(ByVal ws As Worksheet, ByVal dict As Object, _
                         ByVal startRow As Long, ByVal startCol As Long, ByVal topN As Long)
    ws.Cells(startRow, startCol).Value = "名称"
    ws.Cells(startRow, startCol + 1).Value = "売上"
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 1)).Font.Bold = True

    Dim keys As Variant
    keys = KeysByValueDesc(dict)
    Dim n As Long, i As Long
    n = UBound(keys) - LBound(keys) + 1
    If topN > 0 And topN < n Then n = topN

    For i = 0 To n - 1
        ws.Cells(startRow + 1 + i, startCol).Value = keys(i)
        ws.Cells(startRow + 1 + i, startCol + 1).Value = dict(keys(i))
    Next i
    ws.Range(ws.Cells(startRow + 1, startCol + 1), ws.Cells(startRow + n, startCol + 1)) _
      .NumberFormat = "#,##0"
    AddDataBars ws.Range(ws.Cells(startRow + 1, startCol + 1), ws.Cells(startRow + n, startCol + 1))
End Sub

' ---------- 整形ヘルパー ----------
Private Sub AddDataBars(ByVal rng As Range)
    rng.FormatConditions.Delete
    Dim db As Databar
    Set db = rng.FormatConditions.AddDatabar
    db.BarColor.Color = RGB(15, 118, 110)
End Sub

Private Sub MarkNegatives(ByVal rng As Range)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    fc.Font.Color = RGB(200, 30, 30)
    fc.Font.Bold = True
End Sub

' ---------- 並べ替えヘルパー ----------
Private Function SortedKeys(ByVal dict As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    keys = dict.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
    SortedKeys = keys
End Function

Private Function KeysByValueDesc(ByVal dict As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    keys = dict.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If dict(keys(j)) > dict(keys(i)) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
    KeysByValueDesc = keys
End Function
