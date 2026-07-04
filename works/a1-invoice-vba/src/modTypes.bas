Attribute VB_Name = "modTypes"
'==============================================================
' modTypes ― 共有型定義
'==============================================================
Option Explicit

Public Type TConfig
    CompanyName As String       ' 自社名
    CompanyAddress As String    ' 自社住所
    CompanyTel As String        ' 自社TEL
    BankInfo As String          ' 振込先
    RegistrationNo As String    ' 適格請求書発行事業者 登録番号
    TargetMonth As Date         ' 対象月(月初日で保持)
    OutputDir As String         ' PDF出力先ルート
    DueDateRule As Long         ' 支払期限: 翌月末=1 / 翌々月10日=2
    RoundingMode As Long        ' 消費税端数: 0=切り捨て 1=四捨五入 2=切り上げ
    WithholdingEnabled As Boolean ' 源泉徴収の控除欄を出すか
End Type

Public Type TSalesLine
    SalesDate As Date
    Description As String
    Qty As Double
    UnitPrice As Currency
    TaxRate As Double           ' 0.1 / 0.08
End Type

Public Type TRunResult
    SuccessCount As Long
    SkipCount As Long
End Type

' 1顧客分の請求データ(Collection格納用にクラスの代わりの辞書を使う)
' キー: "Code","Name","Honorific","Zip","Address","Withholding"(Boolean),"Lines"(Collection of Variant配列)
Public Function NewLine(ByVal d As Date, ByVal desc As String, ByVal qty As Double, _
                        ByVal unitPrice As Currency, ByVal taxRate As Double) As Variant
    NewLine = Array(d, desc, qty, unitPrice, taxRate)
End Function
