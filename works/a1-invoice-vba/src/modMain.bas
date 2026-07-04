Attribute VB_Name = "modMain"
'==============================================================
' modMain ― エントリポイント
' 請求書自動発行システム InvoiceForge v1.0
' 顧客マスタ×売上明細から対象月の請求書を全顧客分一括生成し、
' 顧客別フォルダへPDF出力・発行履歴を記録する。
'==============================================================
Option Explicit

' 一括発行(リボン/ボタンから呼ぶ)
Public Sub GenerateAllInvoices()
    Dim cfg As TConfig
    Dim customers As Collection
    Dim result As TRunResult
    Dim t0 As Single

    On Error GoTo Fail
    t0 = Timer

    cfg = modConfig.LoadConfig()
    If Not modValidate.ValidateAll(cfg) Then
        MsgBox "入力データにエラーがあります。「検査結果」シートを確認してください。", vbExclamation
        Exit Sub
    End If

    Set customers = modData.LoadCustomersWithSales(cfg)
    If customers.Count = 0 Then
        MsgBox "対象月 " & Format$(cfg.TargetMonth, "yyyy/mm") & " の売上明細がありません。", vbInformation
        Exit Sub
    End If

    result = modInvoice.GenerateBatch(customers, cfg)

    MsgBox "発行完了: " & result.SuccessCount & " 件" & vbCrLf & _
           "スキップ: " & result.SkipCount & " 件(明細0件)" & vbCrLf & _
           "処理時間: " & Format$(Timer - t0, "0.0") & " 秒" & vbCrLf & vbCrLf & _
           "出力先: " & cfg.OutputDir, vbInformation, "InvoiceForge"
    Exit Sub

Fail:
    modLog.LogError "GenerateAllInvoices", Err.Number, Err.Description
    MsgBox "処理を中断しました。" & vbCrLf & _
           "詳細はエラーログシートを確認してください。" & vbCrLf & _
           "(" & Err.Description & ")", vbCritical
End Sub

' 選択顧客のみ発行(顧客マスタ上で行を選択して実行)
Public Sub GenerateSelectedInvoice()
    Dim cfg As TConfig
    Dim code As String

    On Error GoTo Fail
    cfg = modConfig.LoadConfig()

    code = modData.SelectedCustomerCode()
    If Len(code) = 0 Then
        MsgBox "顧客マスタで発行したい顧客の行を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Dim customers As Collection
    Set customers = modData.LoadCustomersWithSales(cfg, code)
    If customers.Count = 0 Then
        MsgBox "顧客 " & code & " の対象月明細がありません。", vbInformation
        Exit Sub
    End If

    Dim result As TRunResult
    result = modInvoice.GenerateBatch(customers, cfg)
    MsgBox "発行完了: " & result.SuccessCount & " 件", vbInformation, "InvoiceForge"
    Exit Sub

Fail:
    modLog.LogError "GenerateSelectedInvoice", Err.Number, Err.Description
    MsgBox "処理を中断しました: " & Err.Description, vbCritical
End Sub
