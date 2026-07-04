Attribute VB_Name = "modConfig"
'==============================================================
' modConfig ― 「設定」シートの読み込み
' シート構成(B列に値): 自社名/住所/TEL/振込先/登録番号/対象月/
'   出力先フォルダ/支払期限ルール/端数処理/源泉徴収欄
'==============================================================
Option Explicit

Public Const SHEET_CONFIG As String = "設定"
Public Const SHEET_CUSTOMER As String = "顧客マスタ"
Public Const SHEET_SALES As String = "売上明細"
Public Const SHEET_TEMPLATE As String = "請求書ひな形"
Public Const SHEET_HISTORY As String = "発行履歴"
Public Const SHEET_ERRORS As String = "エラーログ"
Public Const SHEET_CHECK As String = "検査結果"

Public Function LoadConfig() As TConfig
    Dim ws As Worksheet
    Dim cfg As TConfig

    Set ws = SheetByName(SHEET_CONFIG)

    cfg.CompanyName = ReadRequired(ws, "自社名")
    cfg.CompanyAddress = ReadRequired(ws, "自社住所")
    cfg.CompanyTel = ReadValue(ws, "自社TEL")
    cfg.BankInfo = ReadRequired(ws, "振込先")
    cfg.RegistrationNo = ReadValue(ws, "登録番号")

    Dim tm As Variant
    tm = ReadRequired(ws, "対象月")
    If Not IsDate(tm) Then Err.Raise vbObjectError + 101, , "設定「対象月」が日付ではありません: " & tm
    cfg.TargetMonth = DateSerial(Year(CDate(tm)), Month(CDate(tm)), 1)

    cfg.OutputDir = ReadRequired(ws, "出力先フォルダ")
    If Right$(cfg.OutputDir, 1) <> Application.PathSeparator Then
        cfg.OutputDir = cfg.OutputDir & Application.PathSeparator
    End If

    cfg.DueDateRule = CLng(Val(ReadValue(ws, "支払期限ルール", "1")))
    cfg.RoundingMode = CLng(Val(ReadValue(ws, "端数処理", "0")))
    cfg.WithholdingEnabled = (ReadValue(ws, "源泉徴収欄", "しない") = "する")

    LoadConfig = cfg
End Function

' 設定シートのA列をキーに、同じ行のB列を返す
Private Function ReadValue(ByVal ws As Worksheet, ByVal key As String, _
                           Optional ByVal defaultValue As String = "") As String
    Dim hit As Range
    Set hit = ws.Columns(1).Find(What:=key, LookAt:=xlWhole, MatchCase:=True)
    If hit Is Nothing Then
        ReadValue = defaultValue
    Else
        ReadValue = Trim$(CStr(hit.Offset(0, 1).Value))
        If Len(ReadValue) = 0 Then ReadValue = defaultValue
    End If
End Function

Private Function ReadRequired(ByVal ws As Worksheet, ByVal key As String) As String
    ReadRequired = ReadValue(ws, key)
    If Len(ReadRequired) = 0 Then
        Err.Raise vbObjectError + 100, , "設定シートに「" & key & "」がありません(A列にキー・B列に値)"
    End If
End Function

Public Function SheetByName(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If SheetByName Is Nothing Then
        Err.Raise vbObjectError + 102, , "必要なシート「" & name & "」が見つかりません"
    End If
End Function
