Attribute VB_Name = "modMain"
'==============================================================
' modMain ― エントリポイント
' ReportHub v1.0 ― 複数ブック自動集計ダッシュボード
' 指定フォルダ内の月次報告ブック(支店/担当者別)をすべて読み込み、
' 様式ゆらぎを吸収して統合し、ダッシュボードを自動更新する。
'==============================================================
Option Explicit

Public Sub RefreshDashboard()
    Dim t0 As Single
    t0 = Timer

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim folderPath As String
    folderPath = modUtil.ConfigValue("報告書フォルダ")
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 200, , "報告書フォルダが見つかりません: " & folderPath
    End If

    ' 1) 全ブックを走査して統合(元ファイルは読み取り専用・無変更)
    Dim rows As Collection, report As Collection
    Set report = New Collection
    Set rows = modScan.CollectAll(folderPath, report)

    ' 2) 統合データシートへ書き出し
    modScan.WriteUnified rows

    ' 3) ダッシュボード再構築
    modDash.BuildDashboard rows

    ' 4) 取り込み結果レポート
    modScan.WriteReport report

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "更新完了" & vbCrLf & _
           "取り込み: " & report.Count & " ファイル / 統合 " & rows.Count & " 行" & vbCrLf & _
           "処理時間: " & Format$(Timer - t0, "0.0") & " 秒", vbInformation, "ReportHub"
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    modUtil.LogError "RefreshDashboard", Err.Number, Err.Description
    MsgBox "処理を中断しました: " & Err.Description & vbCrLf & _
           "詳細はエラーログシートを確認してください。", vbCritical
End Sub
