Attribute VB_Name = "modLog"
'==============================================================
' modLog ― 発行履歴・採番・エラーログ
' 発行履歴: A:発行日時 B:請求番号 C:顧客コード D:顧客名 E:合計 F:対象月
' エラーログ: A:日時 B:処理 C:番号 D:内容
'==============================================================
Option Explicit

' 対象月内の次の連番(発行履歴の同月最大値+1)
Public Function NextSequence(ByVal targetMonth As Date) As Long
    Dim ws As Worksheet
    Set ws = EnsureSheet(modConfig.SHEET_HISTORY, _
        Array("発行日時", "請求番号", "顧客コード", "顧客名", "合計", "対象月"))

    Dim lastRow As Long, i As Long, maxSeq As Long
    Dim prefix As String
    prefix = Format$(targetMonth, "yyyymm") & "-"
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    For i = 2 To lastRow
        Dim no As String
        no = CStr(ws.Cells(i, 2).Value)
        If Left$(no, Len(prefix)) = prefix Then
            Dim n As Long
            n = CLng(Val(Mid$(no, Len(prefix) + 1)))
            If n > maxSeq Then maxSeq = n
        End If
    Next i
    NextSequence = maxSeq + 1
End Function

Public Sub LogIssue(ByVal invoiceNo As String, ByVal code As String, ByVal custName As String, _
                    ByVal total As Currency, ByVal targetMonth As Date)
    Dim ws As Worksheet
    Set ws = EnsureSheet(modConfig.SHEET_HISTORY, _
        Array("発行日時", "請求番号", "顧客コード", "顧客名", "合計", "対象月"))
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = invoiceNo
    ws.Cells(r, 3).Value = code
    ws.Cells(r, 4).Value = custName
    ws.Cells(r, 5).Value = total
    ws.Cells(r, 6).Value = Format$(targetMonth, "yyyy/mm")
End Sub

Public Sub LogError(ByVal proc As String, ByVal errNo As Long, ByVal msg As String)
    On Error Resume Next   ' ログ失敗で二次エラーを起こさない
    Dim ws As Worksheet
    Set ws = EnsureSheet(modConfig.SHEET_ERRORS, Array("日時", "処理", "番号", "内容"))
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = proc
    ws.Cells(r, 3).Value = errNo
    ws.Cells(r, 4).Value = msg
End Sub

' シートがなければヘッダーつきで作成して返す
Public Function EnsureSheet(ByVal name As String, ByVal headers As Variant) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            EnsureSheet.Cells(1, i - LBound(headers) + 1).Value = headers(i)
        Next i
        EnsureSheet.Rows(1).Font.Bold = True
    End If
End Function
