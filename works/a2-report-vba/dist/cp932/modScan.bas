Attribute VB_Name = "modScan"
'==============================================================
' modScan ― フォルダ走査・様式ゆらぎ吸収・統合
' 報告ブックの想定: 1シート目に月次明細表(どこかの行にヘッダー)。
' 必要列: 日付 / 担当者 / 商品(または品目・商品名) / 売上(または金額・売上高)
' 列順・ヘッダー行位置・余分な前置行が違っても、ヘッダー探索で吸収する。
'==============================================================
Option Explicit

Private Const MAX_HEADER_SCAN As Long = 10   ' ヘッダーを探す最大行数

' フォルダ内の全 .xlsx/.xls を読み、統合行(Variant配列: 日付,担当,商品,売上,元ファイル)を返す
' report には Array(ファイル名, 状態, 詳細) を積む
Public Function CollectAll(ByVal folderPath As String, ByVal report As Collection) As Collection
    Dim out As New Collection
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(folderPath, 1) <> sep Then folderPath = folderPath & sep

    Dim f As String
    f = Dir$(folderPath & "*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And f <> ThisWorkbook.Name Then   ' 一時ファイル・自分自身を除外
            CollectOne folderPath & f, f, out, report
        End If
        f = Dir$()
    Loop

    Set CollectAll = out
End Function

Private Sub CollectOne(ByVal fullPath As String, ByVal fileName As String, _
                       ByVal out As Collection, ByVal report As Collection)
    Dim wb As Workbook
    On Error GoTo Fail

    Set wb = Workbooks.Open(fileName:=fullPath, ReadOnly:=True, UpdateLinks:=0)

    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)

    ' --- ヘッダー行と列位置を探索(様式ゆらぎ吸収の核) ---
    Dim headerRow As Long, colDate As Long, colStaff As Long, colItem As Long, colSales As Long
    If Not FindHeader(ws, headerRow, colDate, colStaff, colItem, colSales) Then
        report.Add Array(fileName, "警告", "ヘッダー行が特定できずスキップ(日付/担当者/商品/売上 の列が必要)")
        GoTo CleanUp
    End If

    ' --- データ本体を配列で読む ---
    Dim lastRow As Long, r As Long, added As Long, skipped As Long
    lastRow = ws.Cells(ws.Rows.Count, colDate).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        Dim vDate As Variant, vStaff As String, vItem As String, vSales As Variant
        vDate = ws.Cells(r, colDate).Value
        vStaff = Trim$(CStr(ws.Cells(r, colStaff).Value))
        vItem = Trim$(CStr(ws.Cells(r, colItem).Value))
        vSales = ws.Cells(r, colSales).Value

        If IsDate(vDate) And Len(vStaff) > 0 And IsNumeric(vSales) Then
            out.Add Array(CDate(vDate), vStaff, vItem, CDbl(vSales), fileName)
            added = added + 1
        ElseIf Not (IsEmpty(vDate) And Len(vStaff) = 0) Then
            skipped = skipped + 1   ' 小計行・メモ行など
        End If
    Next r

    report.Add Array(fileName, "成功", added & " 行取り込み" & _
                     IIf(skipped > 0, " / " & skipped & " 行を除外(小計・空欄等)", ""))

CleanUp:
    wb.Close SaveChanges:=False
    Exit Sub

Fail:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    report.Add Array(fileName, "失敗", Err.Description)
End Sub

' ヘッダー行を上から探索し、別名(品目/商品名/金額/売上高…)も許容して列位置を返す
Private Function FindHeader(ByVal ws As Worksheet, ByRef headerRow As Long, _
                            ByRef colDate As Long, ByRef colStaff As Long, _
                            ByRef colItem As Long, ByRef colSales As Long) As Boolean
    Dim r As Long, c As Long, lastCol As Long
    For r = 1 To MAX_HEADER_SCAN
        colDate = 0: colStaff = 0: colItem = 0: colSales = 0
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            Select Case NormalizeHeader(CStr(ws.Cells(r, c).Value))
                Case "日付", "売上日", "取引日": colDate = c
                Case "担当者", "担当", "担当者名": colStaff = c
                Case "商品", "品目", "商品名", "サービス": colItem = c
                Case "売上", "金額", "売上高", "売上金額": colSales = c
            End Select
        Next c
        If colDate > 0 And colStaff > 0 And colItem > 0 And colSales > 0 Then
            headerRow = r
            FindHeader = True
            Exit Function
        End If
    Next r
End Function

' 空白除去・全角括弧注記の除去(例: "売上(円)" → "売上")
Private Function NormalizeHeader(ByVal s As String) As String
    s = Replace$(Replace$(Trim$(s), " ", ""), "　", "")
    Dim p As Long
    p = InStr(s, "(")
    If p = 0 Then p = InStr(s, "(")
    If p > 0 Then s = Left$(s, p - 1)
    NormalizeHeader = s
End Function

' 統合データシートへ書き出し
Public Sub WriteUnified(ByVal rows As Collection)
    Dim ws As Worksheet
    Set ws = modUtil.EnsureSheet("統合データ", Array("日付", "担当者", "商品", "売上", "元ファイル"))
    ws.Range("A2:E" & ws.Rows.Count).ClearContents

    If rows.Count = 0 Then Exit Sub
    Dim arr() As Variant, i As Long, item As Variant
    ReDim arr(1 To rows.Count, 1 To 5)
    i = 1
    For Each item In rows
        arr(i, 1) = item(0): arr(i, 2) = item(1): arr(i, 3) = item(2)
        arr(i, 4) = item(3): arr(i, 5) = item(4)
        i = i + 1
    Next item
    ws.Range("A2").Resize(rows.Count, 5).Value = arr   ' 一括書き込み(高速化)
    ws.Columns("A").NumberFormat = "yyyy/mm/dd"
    ws.Columns("D").NumberFormat = "#,##0"
End Sub

' 取り込み結果レポート
Public Sub WriteReport(ByVal report As Collection)
    Dim ws As Worksheet
    Set ws = modUtil.EnsureSheet("取込結果", Array("ファイル", "状態", "詳細"))
    ws.Range("A2:C" & ws.Rows.Count).ClearContents

    Dim i As Long, item As Variant
    i = 2
    For Each item In report
        ws.Cells(i, 1).Value = item(0)
        ws.Cells(i, 2).Value = item(1)
        ws.Cells(i, 3).Value = item(2)
        If item(1) = "失敗" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).Interior.Color = RGB(255, 205, 205)
        ElseIf item(1) = "警告" Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).Interior.Color = RGB(255, 240, 190)
        End If
        i = i + 1
    Next item
End Sub
