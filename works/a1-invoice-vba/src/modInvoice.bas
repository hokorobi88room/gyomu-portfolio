Attribute VB_Name = "modInvoice"
'==============================================================
' modInvoice ― 請求書の生成(ひな形シートへの差し込み)と税計算
' 請求書ひな形シートの名前付きセルへ値を差し込む方式。
' 名前: 宛先/宛先住所/請求番号/発行日/支払期限/自社情報/登録番号/
'       振込先/小計10/消費税10/小計8/消費税8/源泉徴収/合計
' 明細行: セル名「明細開始」から下へ最大 MAX_LINES 行
'==============================================================
Option Explicit

Private Const MAX_LINES As Long = 20   ' ひな形の明細行数(超過分は2枚目へ)

Public Function GenerateBatch(ByVal customers As Collection, ByRef cfg As TConfig) As TRunResult
    Dim result As TRunResult
    Dim cust As Object
    Dim seq As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    seq = modLog.NextSequence(cfg.TargetMonth)

    For Each cust In customers
        If cust("Lines").Count = 0 Then
            result.SkipCount = result.SkipCount + 1
        Else
            Dim invoiceNo As String
            invoiceNo = Format$(cfg.TargetMonth, "yyyymm") & "-" & Format$(seq, "000")
            FillTemplate cust, cfg, invoiceNo
            modPdf.ExportInvoicePdf cust, cfg, invoiceNo
            modLog.LogIssue invoiceNo, CStr(cust("Code")), CStr(cust("Name")), _
                            LastTotal, cfg.TargetMonth
            seq = seq + 1
            result.SuccessCount = result.SuccessCount + 1
        End If
    Next cust

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    GenerateBatch = result
    Exit Function

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description   ' 上位(modMain)で処理
End Function

' 直近に生成した請求書の合計(履歴記録用)
Private mLastTotal As Currency
Public Property Get LastTotal() As Currency
    LastTotal = mLastTotal
End Property

Private Sub FillTemplate(ByVal cust As Object, ByRef cfg As TConfig, ByVal invoiceNo As String)
    Dim ws As Worksheet
    Set ws = modConfig.SheetByName(modConfig.SHEET_TEMPLATE)

    ' --- 明細を先にクリア ---
    Dim startCell As Range
    Set startCell = ws.Range("明細開始")
    startCell.Resize(MAX_LINES, 5).ClearContents

    ' --- 集計(税率別) ---
    Dim sub10 As Currency, sub8 As Currency
    Dim ln As Variant, i As Long
    i = 0
    For Each ln In cust("Lines")
        If i >= MAX_LINES Then
            Err.Raise vbObjectError + 120, , _
                "明細が " & MAX_LINES & " 行を超えました(顧客: " & cust("Name") & _
                ")。ひな形の行数を拡張するか、明細を分割してください。"
        End If
        Dim amount As Currency
        amount = CCur(ln(2) * ln(3))    ' 数量×単価
        startCell.Offset(i, 0).Value = Format$(ln(0), "mm/dd")
        startCell.Offset(i, 1).Value = ln(1)
        startCell.Offset(i, 2).Value = ln(2)
        startCell.Offset(i, 3).Value = ln(3)
        startCell.Offset(i, 4).Value = amount
        If ln(4) = 0.08 Then
            sub8 = sub8 + amount
        Else
            sub10 = sub10 + amount
        End If
        i = i + 1
    Next ln

    ' --- 税計算(税率ごとに端数処理: 適格請求書の要件) ---
    Dim tax10 As Currency, tax8 As Currency
    tax10 = RoundTax(sub10 * 0.1, cfg.RoundingMode)
    tax8 = RoundTax(sub8 * 0.08, cfg.RoundingMode)

    Dim withholding As Currency
    withholding = 0
    If cfg.WithholdingEnabled And CBool(cust("Withholding")) Then
        ' 源泉徴収(報酬・料金 10.21%、100万円以下想定。超過分は要個別対応)
        withholding = Fix((sub10 + sub8) * 0.1021)
    End If

    mLastTotal = sub10 + tax10 + sub8 + tax8 - withholding

    ' --- 差し込み ---
    ws.Range("宛先").Value = cust("Name") & " " & cust("Honorific")
    ws.Range("宛先住所").Value = "〒" & cust("Zip") & vbLf & cust("Address")
    ws.Range("請求番号").Value = invoiceNo
    ws.Range("発行日").Value = Date
    ws.Range("支払期限").Value = DueDate(cfg)
    ws.Range("自社情報").Value = cfg.CompanyName & vbLf & cfg.CompanyAddress & _
                                 IIf(Len(cfg.CompanyTel) > 0, vbLf & "TEL: " & cfg.CompanyTel, "")
    ws.Range("登録番号").Value = IIf(Len(cfg.RegistrationNo) > 0, "登録番号: " & cfg.RegistrationNo, "")
    ws.Range("振込先").Value = cfg.BankInfo
    ws.Range("小計10").Value = sub10
    ws.Range("消費税10").Value = tax10
    ws.Range("小計8").Value = sub8
    ws.Range("消費税8").Value = tax8
    ws.Range("源泉徴収").Value = -withholding
    ws.Range("合計").Value = mLastTotal
End Sub

Private Function RoundTax(ByVal v As Double, ByVal mode As Long) As Currency
    Select Case mode
        Case 1: RoundTax = CCur(WorksheetFunction.Round(v, 0))     ' 四捨五入
        Case 2: RoundTax = CCur(-Int(-v))                          ' 切り上げ
        Case Else: RoundTax = CCur(Fix(v))                         ' 切り捨て(既定)
    End Select
End Function

Private Function DueDate(ByRef cfg As TConfig) As Date
    Select Case cfg.DueDateRule
        Case 2  ' 翌々月10日
            DueDate = DateSerial(Year(cfg.TargetMonth), Month(cfg.TargetMonth) + 2, 10)
        Case Else ' 翌月末
            DueDate = DateSerial(Year(cfg.TargetMonth), Month(cfg.TargetMonth) + 2, 0)
    End Select
End Function
