Attribute VB_Name = "modUtil"
'==============================================================
' modUtil ― 共通ヘルパー(設定・シート・ログ)
'==============================================================
Option Explicit

' 「設定」シートのA列キー→B列値
Public Function ConfigValue(ByVal key As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 201, , "「設定」シートがありません"

    Dim hit As Range
    Set hit = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If hit Is Nothing Then Err.Raise vbObjectError + 202, , "設定「" & key & "」がありません"
    ConfigValue = Trim$(CStr(hit.Offset(0, 1).Value))
End Function

' シートがなければヘッダーつきで作成
Public Function EnsureSheet(ByVal name As String, ByVal headers As Variant) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
    If IsArray(headers) Then
        Dim n As Long
        On Error Resume Next
        n = UBound(headers) - LBound(headers) + 1
        On Error GoTo 0
        If n > 0 And Len(CStr(EnsureSheet.Cells(1, 1).Value)) = 0 Then
            Dim i As Long
            For i = LBound(headers) To UBound(headers)
                EnsureSheet.Cells(1, i - LBound(headers) + 1).Value = headers(i)
            Next i
            EnsureSheet.Rows(1).Font.Bold = True
        End If
    End If
End Function

' 辞書の数値取得(未登録なら0)
Public Function DictNum(ByVal dict As Object, ByVal key As String) As Double
    If dict.Exists(key) Then DictNum = CDbl(dict(key))
End Function

Public Sub LogError(ByVal proc As String, ByVal errNo As Long, ByVal msg As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureSheet("エラーログ", Array("日時", "処理", "番号", "内容"))
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = proc
    ws.Cells(r, 3).Value = errNo
    ws.Cells(r, 4).Value = msg
End Sub
