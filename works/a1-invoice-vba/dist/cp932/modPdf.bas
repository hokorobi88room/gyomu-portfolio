Attribute VB_Name = "modPdf"
'==============================================================
' modPdf ― PDF出力(顧客別フォルダへ振り分け)
' 出力例: <出力先>/2026-06/C001_株式会社サンプル/202606-001_請求書.pdf
'==============================================================
Option Explicit

Public Sub ExportInvoicePdf(ByVal cust As Object, ByRef cfg As TConfig, ByVal invoiceNo As String)
    Dim dirPath As String
    dirPath = cfg.OutputDir & Format$(cfg.TargetMonth, "yyyy-mm") & Application.PathSeparator & _
              SafeName(CStr(cust("Code")) & "_" & CStr(cust("Name")))
    EnsureDir dirPath

    Dim filePath As String
    filePath = dirPath & Application.PathSeparator & invoiceNo & "_請求書.pdf"

    Dim ws As Worksheet
    Set ws = modConfig.SheetByName(modConfig.SHEET_TEMPLATE)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
                           Quality:=xlQualityStandard, OpenAfterPublish:=False
End Sub

' 多階層のフォルダを順に作成(MkDirは1階層ずつしか作れない)
Private Sub EnsureDir(ByVal path As String)
    Dim parts() As String
    Dim built As String
    Dim i As Long
    Dim sep As String
    sep = Application.PathSeparator

    parts = Split(path, sep)
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Len(built) = 0 And InStr(parts(i), ":") > 0 Then
                built = parts(i)            ' "C:" ドライブ部
            Else
                built = built & sep & parts(i)
            End If
            If Len(built) > 2 And Len(Dir$(built, vbDirectory)) = 0 Then
                MkDir built
            End If
        End If
    Next i
End Sub

' ファイル名に使えない文字を全角へ置換
Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim rep As Variant
    rep = Array("￥", "／", "：", "＊", "？", "”", "＜", "＞", "｜")
    SafeName = s
    For i = LBound(bad) To UBound(bad)
        SafeName = Replace$(SafeName, CStr(bad(i)), CStr(rep(i)))
    Next i
End Function
