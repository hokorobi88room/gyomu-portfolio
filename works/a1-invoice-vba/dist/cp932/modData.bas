Attribute VB_Name = "modData"
'==============================================================
' modData ― 顧客マスタ・売上明細の読み込み
' 顧客マスタ: A:顧客コード B:顧客名 C:敬称 D:郵便番号 E:住所 F:源泉(する/しない)
' 売上明細:   A:日付 B:顧客コード C:品目・作業内容 D:数量 E:単価 F:税率(10/8)
' いずれも1行目はヘッダー。高速化のため配列へ一括読込して処理する。
'==============================================================
Option Explicit

' 対象月の売上がある顧客を、明細つきの辞書のCollectionで返す
' onlyCode を指定するとその顧客のみ
Public Function LoadCustomersWithSales(ByRef cfg As TConfig, _
                                       Optional ByVal onlyCode As String = "") As Collection
    Dim out As New Collection
    Dim salesByCode As Object   ' Scripting.Dictionary(遅延バインディング: 参照設定不要)
    Set salesByCode = CreateObject("Scripting.Dictionary")

    ' --- 売上明細を一括読込 ---
    Dim wsS As Worksheet, arr As Variant, i As Long
    Set wsS = modConfig.SheetByName(modConfig.SHEET_SALES)
    arr = UsedBody(wsS)
    If IsEmpty(arr) Then
        Set LoadCustomersWithSales = out
        Exit Function
    End If

    Dim monthFrom As Date, monthTo As Date
    monthFrom = cfg.TargetMonth
    monthTo = DateAdd("m", 1, cfg.TargetMonth)

    For i = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(i, 1)) Then
            Dim d As Date, code As String
            d = CDate(arr(i, 1))
            code = Trim$(CStr(arr(i, 2)))
            If d >= monthFrom And d < monthTo Then
                If Len(onlyCode) = 0 Or code = onlyCode Then
                    If Not salesByCode.Exists(code) Then salesByCode.Add code, New Collection
                    Dim rate As Double
                    rate = Val(arr(i, 6)) / 100#   ' 10 → 0.1
                    salesByCode(code).Add modTypes.NewLine(d, CStr(arr(i, 3)), _
                        CDbl(arr(i, 4)), CCur(arr(i, 5)), rate)
                End If
            End If
        End If
    Next i

    ' --- 顧客マスタと突合 ---
    Dim wsC As Worksheet, arrC As Variant
    Set wsC = modConfig.SheetByName(modConfig.SHEET_CUSTOMER)
    arrC = UsedBody(wsC)
    If IsEmpty(arrC) Then Err.Raise vbObjectError + 110, , "顧客マスタが空です"

    For i = 1 To UBound(arrC, 1)
        Dim c As String
        c = Trim$(CStr(arrC(i, 1)))
        If Len(c) > 0 And salesByCode.Exists(c) Then
            Dim cust As Object
            Set cust = CreateObject("Scripting.Dictionary")
            cust("Code") = c
            cust("Name") = CStr(arrC(i, 2))
            cust("Honorific") = IIf(Len(CStr(arrC(i, 3))) > 0, CStr(arrC(i, 3)), "御中")
            cust("Zip") = CStr(arrC(i, 4))
            cust("Address") = CStr(arrC(i, 5))
            cust("Withholding") = (CStr(arrC(i, 6)) = "する")
            Set cust("Lines") = salesByCode(c)
            out.Add cust
            salesByCode.Remove c
        End If
    Next i

    ' マスタに存在しない顧客コードの明細が残っていたら異常(検収条件: 黙って捨てない)
    If salesByCode.Count > 0 Then
        Dim k As Variant, missing As String
        For Each k In salesByCode.Keys
            missing = missing & k & " "
        Next k
        Err.Raise vbObjectError + 111, , _
            "売上明細に、顧客マスタに存在しない顧客コードがあります: " & missing
    End If

    Set LoadCustomersWithSales = out
End Function

' 選択中セルの行から顧客コードを取得(顧客マスタ上でのみ有効)
Public Function SelectedCustomerCode() As String
    If ActiveSheet.Name <> modConfig.SHEET_CUSTOMER Then Exit Function
    Dim r As Long
    r = Selection.Cells(1, 1).Row
    If r < 2 Then Exit Function
    SelectedCustomerCode = Trim$(CStr(ActiveSheet.Cells(r, 1).Value))
End Function

' ヘッダー行を除いた使用範囲を2次元配列で返す(空なら Empty)
Public Function UsedBody(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Function
    UsedBody = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value2
    ' 1行だけのとき Value2 はスカラー配列にならないため正規化
    If Not IsArray(UsedBody) Then
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = UsedBody
        UsedBody = tmp
    End If
End Function
