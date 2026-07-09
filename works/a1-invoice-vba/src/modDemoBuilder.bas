Attribute VB_Name = "modDemoBuilder"
'==============================================================
' modDemoBuilder ― 体験版セットアップ & ワンクリック請求書生成
'--------------------------------------------------------------
' このモジュールだけで「開く→ボタンを押す→請求書が全件できる」を
' 完結させる、営業デモ/検収体験用の自己完結ランチャーである。
'   ・BuildDemoWorkbook … サンプルデータ(取引先10社・売上約1,000件)・
'       スタート画面(自社設定つき)・ボタンを一括生成する。
'   ・請求書をまとめて作る … 緑ボタンから呼ばれ、売上明細を取引先
'       ごとに自動集計して請求書を生成する(体験の本番)。
'   ・請求書をPDFで保存する … 生成済み請求書を「社名_年月.pdf」で保存。
' 自社情報は「スタート」シートの黄色い欄(名前定義 Cfg*)から読み取る。
'==============================================================
Option Explicit

Private Const S_START As String = "スタート"
Private Const S_CUST  As String = "顧客マスタ"
Private Const S_SALES As String = "売上明細"
Private Const INK     As Long = &H1A2433      ' 濃紺(文字)
Private Const ACCENT  As Long = &H6E760F      ' ティール(BGR)
Private Const PALE    As Long = &HFAFDF0      ' 薄ティール
Private Const EDITBG  As Long = &HD6FBFF      ' 淡い黄(編集OK欄)  BGR
Private Const EDITBRD As Long = &H8CCDD6      ' 黄枠

Private Const N_CUST  As Long = 10            ' 取引先数(=請求書シート数)

' 自社情報の既定値(スタートの黄色い欄に初期表示 / 未入力時のフォールバック)
Private Const DFLT_NAME As String = "マルヤマ鉄工所 株式会社"
Private Const DFLT_ZIP  As String = "441-8021"
Private Const DFLT_ADDR As String = "愛知県豊橋市白河町1-2-3"
Private Const DFLT_TEL  As String = "0532-00-0000"
Private Const DFLT_REG  As String = "T1234567890123"
Private Const DFLT_BANK As String = "○○銀行 △△支店 普通 1234567  マルヤマテッコウ(カ"

' 自動ビルド/自動検証(COM)からの実行時にMsgBoxを抑止するフラグ。
Private gDemoSilent As Boolean
Private gLastError As String     ' 直近のエラー内容(自動検証がRun後に読み取る。空文字=成功)

Public Sub DemoSetSilent(ByVal silent As Boolean)
    gDemoSilent = silent
End Sub

Public Function DemoLastError() As String
    DemoLastError = gLastError
End Function

'==============================================================
' 組み立て担当が一度だけ実行するセットアップ
'==============================================================
Public Sub BuildDemoWorkbook()
    On Error GoTo Fail
    gLastError = ""

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    RemoveGeneratedInvoices
    BuildCustomerSheet
    BuildSalesSheet
    BuildStartSheet
    RemoveDefaultBlankSheets

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(S_START).Activate
    On Error GoTo Fail

    If Not gDemoSilent Then
        MsgBox "体験版の準備ができました。" & vbCrLf & vbCrLf & _
               "「スタート」シートの緑のボタンを押すと、" & vbCrLf & _
               "取引先ごとの請求書がまとめて出来上がります。", _
               vbInformation, "InvoiceForge 体験版"
    End If
    Exit Sub
Fail:
    gLastError = "BuildDemoWorkbook: " & Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not gDemoSilent Then MsgBox "セットアップでエラー: " & Err.Description, vbCritical
End Sub

'==============================================================
' 【体験の本番】売上明細を取引先ごとに集計して請求書を一括生成
'==============================================================
Public Sub 請求書をまとめて作る()
    Dim wsC As Worksheet, wsS As Worksheet
    Dim t0 As Single, made As Long
    Dim custD As Object, idxD As Object, orderCodes As Collection
    Dim lastC As Long, lastS As Long
    Dim cData As Variant, sData As Variant
    Dim ci As Long, si As Long, ccode As String, scode As String
    Dim code As Variant, rowColl As Collection, cnt As Long
    Dim idxArr() As Long, adBlk() As Variant, rateArr() As Variant
    Dim rowIdx As Variant, kk As Long, m As Long, ri As Long
    Dim rate As Double, hasSub8 As Boolean
    Dim cust As Variant

    On Error GoTo Fail
    gLastError = ""
    t0 = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsC = Sheet(S_CUST)
    Set wsS = Sheet(S_SALES)
    RemoveGeneratedInvoices

    Set custD = CreateObject("Scripting.Dictionary")
    lastC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    If lastC >= 2 Then
        cData = wsC.Range("A2:F" & lastC).Value
        For ci = 1 To UBound(cData, 1)
            ccode = Trim$(CStr(cData(ci, 1)))
            If Len(ccode) > 0 Then
                custD(ccode) = Array( _
                    CStr(cData(ci, 2)), _
                    IIf(Len(CStr(cData(ci, 3))) > 0, CStr(cData(ci, 3)), "御中"), _
                    CStr(cData(ci, 4)), _
                    CStr(cData(ci, 5)), _
                    (CStr(cData(ci, 6)) = "する"))
            End If
        Next ci
    End If

    lastS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastS < 2 Then Err.Raise vbObjectError + 201, , "売上明細が空です"
    sData = wsS.Range("A2:F" & lastS).Value
    Set orderCodes = New Collection
    Set idxD = CreateObject("Scripting.Dictionary")
    For si = 1 To UBound(sData, 1)
        scode = Trim$(CStr(sData(si, 2)))
        If Len(scode) > 0 Then
            If Not idxD.Exists(scode) Then
                idxD.Add scode, New Collection
                orderCodes.Add scode
            End If
            idxD(scode).Add si
        End If
    Next si

    For Each code In orderCodes
        Set rowColl = idxD(CStr(code))
        cnt = rowColl.Count
        If cnt > 0 Then
            ' 明細の行番号を集め、日付昇順(同日は品目名順)に並べ替えてから請求書化する
            ReDim idxArr(1 To cnt)
            kk = 0
            For Each rowIdx In rowColl
                kk = kk + 1: idxArr(kk) = CLng(rowIdx)
            Next rowIdx
            SortRowsByDateName idxArr, sData

            ReDim adBlk(1 To cnt, 1 To 4)    ' A:D  日付/品目/数量/単価
            ReDim rateArr(1 To cnt, 1 To 1)  ' G    税率(SUMIF用・非表示)
            hasSub8 = False
            For m = 1 To cnt
                ri = idxArr(m)
                rate = Val(sData(ri, 6))
                adBlk(m, 1) = sData(ri, 1)
                adBlk(m, 2) = CStr(sData(ri, 3))
                adBlk(m, 3) = CDbl(sData(ri, 4))
                adBlk(m, 4) = CCur(sData(ri, 5))
                rateArr(m, 1) = rate
                If rate = 8 Then hasSub8 = True
            Next m
            If custD.Exists(CStr(code)) Then
                cust = custD(CStr(code))
            Else
                cust = Array("(取引先未登録 " & CStr(code) & ")", "御中", "", "", False)
            End If
            MakeOneInvoice CStr(code), cust, adBlk, rateArr, cnt, hasSub8
            made = made + 1
        End If
    Next code

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error Resume Next
    Sheet(S_START).Activate
    On Error GoTo Fail
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not gDemoSilent Then
        MsgBox made & " 社ぶんの請求書を " & Format$(Timer - t0, "0.0") & " 秒で作成しました。" & vbCrLf & vbCrLf & _
               "売上明細 " & Format$(lastS - 1, "#,##0") & " 件を取引先ごとに自動集計し、" & vbCrLf & _
               "消費税(10%/8%)・源泉徴収まで計算済みです。" & vbCrLf & _
               "画面下の「請求_」タブを開いて確認してください。", _
               vbInformation, "できあがりました"
    End If
    Exit Sub
Fail:
    gLastError = "MakeInvoices: " & Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not gDemoSilent Then MsgBox "請求書の作成でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================
' 生成済み請求書を「社名_年月.pdf」で保存(任意)
'==============================================================
Public Sub 請求書をPDFで保存する()
    Dim ws As Worksheet, n As Long, pdfDir As String
    Dim fn As String, rawName As String, period As String, lastR As Long
    On Error GoTo Fail
    gLastError = ""
    ' 変数名は "dir" にしない。VBA組み込み Dir() と衝突しコンパイルエラーになるため。
    pdfDir = ThisWorkbook.Path & Application.PathSeparator & "請求書PDF"
    If Len(Dir(pdfDir, vbDirectory)) = 0 Then MkDir pdfDir

    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If IsInvoiceSheet(ws) Then
            rawName = CStr(ws.Range("AA1").Value)         ' 取引先名(生)
            period = CStr(ws.Range("AA2").Value)          ' yyyy年MM月
            lastR = Val(ws.Range("AA3").Value)            ' 印刷範囲の最終行
            If lastR < 1 Then lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If Len(rawName) = 0 Then rawName = ws.Name
            If Len(period) = 0 Then period = Format$(Date, "yyyy年MM月")
            ws.PageSetup.PrintArea = "$A$1:$E$" & lastR
            SetInvoicePage ws
            fn = SafeName(rawName & "_" & period & "分") & ".pdf"
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=pdfDir & Application.PathSeparator & fn, _
                Quality:=xlQualityStandard, OpenAfterPublish:=False
            n = n + 1
        End If
    Next ws
    Application.ScreenUpdating = True
    If Not gDemoSilent Then
        If n = 0 Then
            MsgBox "先に「請求書をまとめて作る」を押してください。", vbExclamation
        Else
            MsgBox n & " 件のPDFを保存しました(社名_年月.pdf)。" & vbCrLf & pdfDir, vbInformation
        End If
    End If
    Exit Sub
Fail:
    gLastError = "SavePdf: " & Err.Description
    Application.ScreenUpdating = True
    If Not gDemoSilent Then MsgBox "PDF保存でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================
' 1取引先分の請求書シートを整形して生成(明細は配列で一括書き込み)
'==============================================================
Private Sub MakeOneInvoice(ByVal code As String, ByVal cust As Variant, _
                           ByRef adBlk As Variant, ByRef rateArr As Variant, _
                           ByVal cnt As Long, ByVal hasSub8 As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = UniqueSheetName(cust(0), code)   ' シート名は社名

    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 2

    ' タイトル
    With ws.Range("A1:E1")
        .Merge
        .Value = "請　求　書"
        .Font.Size = 26
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Color = INK
    End With
    ws.Range("A2:E2").Merge
    With ws.Range("A1:E2").Borders(xlEdgeBottom)
        .LineStyle = xlDouble: .Color = INK: .Weight = xlThick
    End With

    ' 対象年月(明細の月から導出)
    Dim periodDisp As String, periodFile As String, firstDate As Date
    firstDate = CDate(adBlk(1, 1))
    periodDisp = Format$(firstDate, "yyyy年m月")
    periodFile = Format$(firstDate, "yyyy年MM月")

    ' 宛先(左)
    ws.Range("A4").Value = cust(0) & "　" & cust(1)
    ws.Range("A4").Font.Size = 15
    ws.Range("A4").Font.Bold = True
    With ws.Range("A4:C4").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = INK
    End With
    ws.Range("A5").Value = "〒" & cust(2)
    ws.Range("A6").Value = cust(3)
    ws.Range("A5:A6").Font.Size = 9

    ' 請求メタ(右)
    Dim issueDate As Date, due As Date
    issueDate = Date
    due = DateSerial(Year(issueDate), Month(issueDate) + 2, 0)   ' 翌月末
    ws.Range("D4").Value = "請求番号"
    ws.Range("E4").Value = Format$(issueDate, "yyyymm") & "-" & code
    ws.Range("D5").Value = "発行日"
    ws.Range("E5").Value = Format$(issueDate, "yyyy/mm/dd")
    ws.Range("D6").Value = "支払期限"
    ws.Range("E6").Value = Format$(due, "yyyy/mm/dd")
    ws.Range("D4:D6").Font.Size = 9
    ws.Range("D4:D6").Font.Color = RGB(90, 107, 128)
    ws.Range("E4:E6").HorizontalAlignment = xlRight

    ws.Range("A8").Value = periodDisp & "分 として、下記のとおりご請求申し上げます。"

    ' 明細ヘッダー(行10)
    Const HR As Long = 10
    WriteHeaderCell ws.Range("A" & HR), "日付"
    WriteHeaderCell ws.Range("B" & HR), "品目・作業内容"
    WriteHeaderCell ws.Range("C" & HR), "数量"
    WriteHeaderCell ws.Range("D" & HR), "単価"
    WriteHeaderCell ws.Range("E" & HR), "金額"

    ' 明細を一括書き込み。金額(E)は数式 =数量*単価。税率は非表示のG列(SUMIF用)。
    Dim r0 As Long, lastLine As Long
    r0 = HR + 1
    lastLine = r0 + cnt - 1
    ws.Range(ws.Cells(r0, 1), ws.Cells(lastLine, 4)).Value = adBlk           ' A:D 日付/品目/数量/単価
    ws.Range("E" & r0 & ":E" & lastLine).Formula = "=C" & r0 & "*D" & r0      ' 金額=数量*単価(各行へ相対展開)
    ws.Range("G" & r0 & ":G" & lastLine).Value = rateArr                      ' 税率(非表示・SUMIF用)
    ws.Columns("G").Hidden = True

    With ws.Range("A" & r0 & ":E" & lastLine)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(213, 220, 230)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = INK
    End With
    ws.Range("A" & r0 & ":A" & lastLine).NumberFormat = "mm/dd"
    ws.Range("A" & r0 & ":A" & lastLine).HorizontalAlignment = xlCenter
    ws.Range("C" & r0 & ":C" & lastLine).NumberFormat = "#,##0"
    ws.Range("D" & r0 & ":E" & lastLine).NumberFormat = "#,##0""円"""

    ' 合計ブロック ― すべて数式。数量/単価を直すと自動で再計算される。
    Dim erng As String, grng As String
    erng = "E" & r0 & ":E" & lastLine
    grng = "G" & r0 & ":G" & lastLine

    ' 内訳の行位置を先に決める(ご請求金額の帯の下に配置)
    Dim br As Long: br = lastLine + 2
    Dim sub10row As Long, tax10row As Long, sub8row As Long, tax8row As Long, whrow As Long
    Dim nextr As Long
    sub10row = br + 1
    tax10row = br + 2
    nextr = br + 3
    If hasSub8 Then
        sub8row = nextr: tax8row = nextr + 1: nextr = nextr + 2
    End If
    If cust(4) Then
        whrow = nextr: nextr = nextr + 1
    End If

    ' (1) ご請求金額(合計)を先頭に。内訳セルを合計する数式にする
    Dim totalF As String
    totalF = "=E" & sub10row & "+E" & tax10row
    If hasSub8 Then totalF = totalF & "+E" & sub8row & "+E" & tax8row
    If cust(4) Then totalF = totalF & "+E" & whrow
    ws.Range("B" & br & ":C" & br).Merge
    ws.Range("B" & br).Value = "ご請求金額 (税込)"
    ws.Range("B" & br).Font.Bold = True
    ws.Range("B" & br).Font.Size = 12
    ws.Range("B" & br).HorizontalAlignment = xlRight
    ws.Range("B" & br).VerticalAlignment = xlCenter
    ws.Range("D" & br & ":E" & br).Merge
    ws.Range("D" & br).Formula = totalF
    ws.Range("D" & br).NumberFormat = "#,##0""円"""
    ws.Range("D" & br).Font.Size = 16
    ws.Range("D" & br).Font.Bold = True
    ws.Range("D" & br).Font.Color = ACCENT
    ws.Range("D" & br).HorizontalAlignment = xlRight
    ws.Range("D" & br).VerticalAlignment = xlCenter
    ws.Range("B" & br & ":E" & br).Interior.Color = PALE
    ws.Range("B" & br & ":E" & br).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=ACCENT
    ws.Rows(br).RowHeight = 26

    ' (2) 右側(D:E)に内訳 ― 数式(小計=SUMIF / 消費税=ROUNDDOWN / 源泉=ROUNDDOWN)
    PutSumF ws, sub10row, "小計(10%対象)", "=SUMIF(" & grng & ",10," & erng & ")"
    PutSumF ws, tax10row, "消費税(10%)", "=ROUNDDOWN(E" & sub10row & "*0.1,0)"
    If hasSub8 Then
        PutSumF ws, sub8row, "小計(8%対象)", "=SUMIF(" & grng & ",8," & erng & ")"
        PutSumF ws, tax8row, "消費税(8%)", "=ROUNDDOWN(E" & sub8row & "*0.08,0)"
    End If
    If cust(4) Then
        Dim whBase As String
        whBase = "E" & sub10row
        If hasSub8 Then whBase = whBase & "+E" & sub8row
        ' 源泉徴収税額は二段階(所得税法): 100万円以下=10.21%、100万円超の部分=20.42%
        PutSumF ws, whrow, "源泉徴収(-)", _
            "=-IF((" & whBase & ")<=1000000,ROUNDDOWN((" & whBase & ")*0.1021,0)," & _
            "ROUNDDOWN(((" & whBase & ")-1000000)*0.2042,0)+102100)"
    End If

    ' (3) 左側(A)に振込先・発行元(自社)。自社設定から読み、1行ずつ別セルにして折り返さない
    Dim la As Long: la = br + 1
    ws.Range("A" & la).Value = "【お振込先】 " & Cfg("CfgBank", DFLT_BANK)
    ws.Range("A" & (la + 2)).Value = Cfg("CfgName", DFLT_NAME)
    ws.Range("A" & (la + 2)).Font.Bold = True
    ws.Range("A" & (la + 3)).Value = "〒" & Cfg("CfgZip", DFLT_ZIP) & " " & Cfg("CfgAddr", DFLT_ADDR)
    ws.Range("A" & (la + 4)).Value = "TEL " & Cfg("CfgTel", DFLT_TEL) & " ／ 登録番号 " & Cfg("CfgReg", DFLT_REG)
    ws.Range("A" & la & ":A" & (la + 4)).Font.Size = 9

    Dim lastFooter As Long
    lastFooter = la + 4
    If nextr - 1 > lastFooter Then lastFooter = nextr - 1

    ' フォントは使用範囲のみ(全セルに掛けると重い)
    ws.Range("A1:F" & lastFooter).Font.Name = "游ゴシック"

    ' PDF/印刷用メタ(印刷範囲外の AA 列。PrintAreaで除外)。
    ' AA2は必ず文字列にする。そうしないと "2026年06月" が日付に自動変換され、
    ' 読み戻すと "2026/06/01" となり、ファイル名に "/" が混ざって保存に失敗する。
    ws.Range("AA1:AA2").NumberFormat = "@"
    ws.Range("AA1").Value = cust(0)
    ws.Range("AA2").Value = periodFile
    ws.Range("AA3").Value = lastFooter
End Sub

' 印刷/PDF体裁(幅は1ページ、縦は明細が長ければ複数ページ)
Private Sub SetInvoicePage(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$10:$10"      ' 明細見出しを各ページの先頭で繰り返す(複数ページ対策)
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.6)
    End With
    On Error GoTo 0
End Sub

'==============================================================
' サンプルデータ: 顧客マスタ(取引先) ― 鉄工所の得意先
'==============================================================
Private Sub BuildCustomerSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(S_CUST)
    ws.Cells.Clear

    Dim nm As Variant, zp As Variant, ad As Variant
    nm = Array( _
        "大和自動車部品株式会社", "サンライズ機械工業株式会社", "北陸精密工業株式会社", "みどり農機サービス", "東海プラント設備株式会社", _
        "富士精工株式会社", "信州テクノ工業株式会社", "三河金属工業株式会社", "近畿産業機械株式会社", "関東製作所株式会社", _
        "浪速螺子製作所", "甲信重工株式会社", "瀬戸内鋼業株式会社", "天竜鉄工株式会社", "越前工機株式会社", _
        "加賀テクニカル株式会社", "出雲精密株式会社", "讃岐製鋼株式会社", "阿波エンジニアリング株式会社", "筑豊工業株式会社", _
        "日向金属加工株式会社", "陸奥鉄工所", "津軽機工株式会社", "房総プラント株式会社", "湖東産業株式会社", _
        "但馬精機株式会社", "熊野テック株式会社", "城東鉄工株式会社", "有明重機株式会社", "みなと溶接工業")
    zp = Array( _
        "446-0051", "532-0011", "920-0344", "329-0611", "455-0032", _
        "417-0001", "390-0852", "444-0840", "577-0013", "332-0004", _
        "550-0014", "400-0862", "721-0942", "431-3314", "915-0083", _
        "922-0257", "693-0001", "761-0102", "770-8054", "820-0001", _
        "883-0068", "031-0072", "036-8004", "290-0056", "522-0038", _
        "668-0831", "519-3612", "120-0034", "849-0918", "231-0001")
    ad = Array( _
        "愛知県安城市今池町2-4-1", "大阪府大阪市淀川区西中島3-8-2", "石川県金沢市畝田東5-6-7", "栃木県河内郡上三川町本郷1-2", "愛知県名古屋市港区入船2-3-4", _
        "静岡県富士市今泉7-1-2", "長野県松本市島立1000", "愛知県岡崎市羽根町2-5", "大阪府東大阪市長田中2-1-3", "埼玉県川口市領家4-2-1", _
        "大阪府大阪市西区北堀江3-2-6", "山梨県甲府市朝気1-3-5", "広島県福山市引野町3-1", "静岡県浜松市天竜区二俣町2-4", "福井県越前市家久町5-8", _
        "石川県加賀市作見町ル25", "島根県出雲市渡橋町1000", "香川県高松市春日町1620", "徳島県徳島市南末広町2-1", "福岡県飯塚市小正2-3", _
        "宮崎県日向市材木町1-5", "青森県八戸市江陽3-2-1", "青森県弘前市神田4-1", "千葉県市原市五井中央東1-2", "滋賀県彦根市西今町1200", _
        "兵庫県豊岡市九日市上町180", "三重県尾鷲市中川2-4", "東京都足立区千住5-3-2", "佐賀県佐賀市兵庫北3-1", "神奈川県横浜市中区新港1-1")

    Dim out() As Variant, i As Long, hon As String, wh As String
    ReDim out(1 To N_CUST + 1, 1 To 6)
    out(1, 1) = "取引先コード": out(1, 2) = "取引先名": out(1, 3) = "敬称"
    out(1, 4) = "郵便番号": out(1, 5) = "住所": out(1, 6) = "源泉"
    For i = 0 To N_CUST - 1
        If InStr(nm(i), "サービス") > 0 Then hon = "様" Else hon = "御中"
        ' 鉄工所の部品加工代は源泉徴収の対象外。サンプルは全社「しない」。
        ' (源泉が必要な業種は、この列を「する」にすれば二段階税率で自動計算される)
        wh = "しない"
        out(i + 2, 1) = 101 + i
        out(i + 2, 2) = nm(i)
        out(i + 2, 3) = hon
        out(i + 2, 4) = zp(i)
        out(i + 2, 5) = ad(i)
        out(i + 2, 6) = wh
    Next i
    ws.Range("A1:F" & (N_CUST + 1)).Value = out
    HeaderRow ws, 1, 6
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 6
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 34
    ws.Columns("F").ColumnWidth = 6
End Sub

'==============================================================
' サンプルデータ: 売上明細(約1,000件) ― 部品加工・溶接など町工場の仕事
'==============================================================
Private Sub BuildSalesSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(S_SALES)
    ws.Cells.Clear

    Dim itName As Variant, itPrice As Variant, itTax As Variant, itQL As Variant, itQH As Variant
    itName = Array( _
        "ブラケット部品 旋盤加工", "シャフト フライス加工", "バリ取り・仕上げ", _
        "治具部品 マシニング加工", "架台 溶接組立", "材料手配 SS400 t9", _
        "寸法検査・検査成績書作成", "表面処理(無電解ニッケル)外注手配", "熱処理(高周波焼入れ)外注手配", _
        "レーザー切断加工", "ベンダー曲げ加工", "タップ立て加工", _
        "配管フランジ 穴あけ加工", "ステンレス架台 製作一式", "アルミ部品 マシニング加工", _
        "歯車 ホブ切り加工", "溶接補修・現地調整", "組立・調整作業", _
        "材料手配 SUS304 t6", "ショットブラスト処理", "塗装(防錆)外注手配", _
        "ワイヤーカット加工", "現地据付・搬入", "図面作成・データ化", _
        "ステンレス手すり 製作", "納品用木枠梱包・運送手配")
    itPrice = Array(480, 1200, 90, 3200, 28000, 46000, 15000, 850, 1100, 320, 260, 70, 650, 74000, 2600, 4200, 38000, 18000, 58000, 240, 900, 5600, 45000, 12000, 32000, 6500)
    itTax = Array(10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 8)
    itQL = Array(50, 10, 100, 5, 1, 1, 1, 20, 20, 50, 50, 100, 30, 1, 5, 2, 1, 1, 1, 50, 20, 1, 1, 1, 1, 1)
    itQH = Array(300, 120, 500, 60, 8, 3, 2, 200, 150, 400, 400, 600, 200, 4, 80, 40, 2, 5, 3, 300, 150, 30, 2, 3, 6, 6)
    Dim nItems As Long: nItems = UBound(itName) + 1

    Dim total As Long, i As Long, k As Long, lc As Long
    total = 0
    For i = 0 To N_CUST - 1
        total = total + LinesFor(i)
    Next i

    Dim out() As Variant: ReDim out(1 To total + 1, 1 To 6)
    out(1, 1) = "日付": out(1, 2) = "取引先コード": out(1, 3) = "品目・作業内容"
    out(1, 4) = "数量": out(1, 5) = "単価": out(1, 6) = "税率"

    Dim rr As Long, it As Long, q As Long, dday As Long
    rr = 1
    For i = 0 To N_CUST - 1
        lc = LinesFor(i)
        For k = 0 To lc - 1
            rr = rr + 1
            it = (i * 7 + k * 13 + (k \ 3)) Mod nItems
            dday = 1 + ((i * 3 + k * 7) Mod 27)
            q = itQL(it) + (((i + k * 3) * 11) Mod (CLng(itQH(it)) - CLng(itQL(it)) + 1))
            out(rr, 1) = DateSerial(2026, 6, dday)
            out(rr, 2) = 101 + i
            out(rr, 3) = itName(it)
            out(rr, 4) = q
            out(rr, 5) = itPrice(it)
            out(rr, 6) = itTax(it)
        Next k
    Next i

    ws.Range("A1:F" & (total + 1)).Value = out
    HeaderRow ws, 1, 6
    ws.Range("A2:A" & (total + 1)).NumberFormat = "yyyy/mm/dd"
    ws.Range("D2:E" & (total + 1)).NumberFormat = "#,##0"
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 8
End Sub

' 取引先ごとの明細件数(約70-130件で変化)。決定的に生成し再現性を保つ。
Private Function LinesFor(ByVal i As Long) As Long
    LinesFor = 70 + ((i * 37 + 13) Mod 61)
End Function

'==============================================================
' スタート画面 + 実行ボタン + 自社設定
'==============================================================
Private Sub BuildStartSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(S_START)
    ws.Cells.Clear
    On Error Resume Next
    ws.Buttons.Delete
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0

    ThisWorkbook.Worksheets(S_START).Move Before:=ThisWorkbook.Worksheets(1)

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 16

    ws.Range("B2").Value = "請求書 自動発行システム  ―  体験版"
    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = ACCENT

    ws.Range("B4").Value = "「売上明細」に入っている約1,000件の取引を、取引先ごとに自動で集計し、"
    ws.Range("B5").Value = "全取引先ぶんの請求書を一瞬でまとめて作ります。"
    ws.Range("B7").Value = "▼ まずは下の緑のボタンを押してみてください"
    ws.Range("B7").Font.Bold = True

    ' 実行ボタン(緑) ― 2つを明確に離し、それぞれ枠が見えるようにする
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("B9").Left, ws.Range("B9").Top, 330, 52)
    btn.OnAction = "請求書をまとめて作る"
    btn.Caption = "→ 請求書をまとめて作る"
    btn.Font.Size = 15
    btn.Font.Bold = True

    ' ボタン①と②の間に区切りラベルを置き、別々のボタンだと分かるようにする
    ws.Range("B13").Value = "─ できあがったら ─"
    ws.Range("B13").Font.Color = RGB(150, 160, 175)
    ws.Range("B13").Font.Size = 9

    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(ws.Range("B15").Left, ws.Range("B15").Top, 330, 40)
    btn2.OnAction = "請求書をPDFで保存する"
    btn2.Caption = "② 請求書をPDFで保存する（社名_年月.pdf）"
    btn2.Font.Size = 11

    ws.Range("B19").Value = "※ 緑ボタンで、約1,000件の売上明細を取引先ごとに自動集計し、"
    ws.Range("B20").Value = "　 全取引先ぶんの請求書を数秒で作成します。消費税(10%/8%)・源泉徴収も自動。"
    ws.Range("B21").Value = "※ サンプルの数字・取引先・社名はすべて架空です。"
    ws.Range("B19:B21").Font.Color = RGB(90, 107, 128)
    ws.Range("B19:B21").Font.Size = 10

    ' ―― 自社情報(設定) ――
    Dim sr As Long: sr = 24
    ws.Range("B" & sr).Value = "■ 自社情報（この黄色い欄を書き換えると、作成する請求書に反映されます）"
    ws.Range("B" & sr).Font.Bold = True
    ws.Range("B" & sr).Font.Color = INK
    AddSetting ws, sr + 1, "自社名", DFLT_NAME, "CfgName"
    AddSetting ws, sr + 2, "郵便番号", DFLT_ZIP, "CfgZip"
    AddSetting ws, sr + 3, "住所", DFLT_ADDR, "CfgAddr"
    AddSetting ws, sr + 4, "電話番号", DFLT_TEL, "CfgTel"
    AddSetting ws, sr + 5, "登録番号", DFLT_REG, "CfgReg"
    AddSetting ws, sr + 6, "振込先", DFLT_BANK, "CfgBank"
End Sub

' 自社設定の1項目(ラベル + 黄色い編集欄 + 名前定義)を作る
Private Sub AddSetting(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                       ByVal dflt As String, ByVal nm As String)
    ws.Range("B" & r).Value = label
    ws.Range("B" & r).Font.Color = RGB(90, 107, 128)
    ws.Range("C" & r & ":F" & r).Merge
    ws.Range("C" & r).Value = dflt
    With ws.Range("C" & r)
        .Interior.Color = EDITBG
        .Borders.LineStyle = xlContinuous
        .Borders.Color = EDITBRD
    End With
    SetName nm, "$C$" & r
End Sub

'============================ ヘルパー ============================

' 自社設定を名前定義から読む(空なら既定値)
Private Function Cfg(ByVal nm As String, ByVal dflt As String) As String
    Dim v As String
    On Error Resume Next
    v = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
    On Error GoTo 0
    If Len(Trim$(v)) = 0 Then Cfg = dflt Else Cfg = v
End Function

' 名前定義を(あれば作り直して)スタートのセルに割り当てる
Private Sub SetName(ByVal nm As String, ByVal cellA1 As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & S_START & "'!" & cellA1
End Sub

' ファイル名に使えない文字を除去
Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, r As String
    r = s
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each b In bad
        r = Replace(r, CStr(b), "_")
    Next b
    SafeName = Trim$(r)
End Function

' 行番号配列を、売上明細の 日付(列1)昇順 → 同日は品目名(列3)昇順 に並べ替える(挿入ソート)
Private Sub SortRowsByDateName(ByRef idx() As Long, ByRef sData As Variant)
    Dim i As Long, j As Long, kv As Long
    For i = LBound(idx) + 1 To UBound(idx)
        kv = idx(i)
        j = i - 1
        Do While j >= LBound(idx)
            If RowLess(kv, idx(j), sData) Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = kv
    Next i
End Sub

' a行が b行より前(日付が早い。同日なら品目名が先)なら True
Private Function RowLess(ByVal a As Long, ByVal b As Long, ByRef sData As Variant) As Boolean
    Dim da As Double, db As Double
    da = CDbl(sData(a, 1)): db = CDbl(sData(b, 1))     ' 日付シリアル値
    If da < db Then RowLess = True: Exit Function
    If da > db Then RowLess = False: Exit Function
    RowLess = (StrComp(CStr(sData(a, 3)), CStr(sData(b, 3)), vbTextCompare) < 0)
End Function

' 生成した請求書シートか(固定3シート以外で、AA1に社名マーカーを持つもの)
Private Function IsInvoiceSheet(ByVal ws As Worksheet) As Boolean
    Dim nm As String: nm = ws.Name
    If nm = S_START Or nm = S_CUST Or nm = S_SALES Then Exit Function
    IsInvoiceSheet = (Len(CStr(ws.Range("AA1").Value)) > 0)
End Function

' シート名に使えない文字を除き31文字以内へ
Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, r As String
    r = s
    bad = Array("\", "/", "?", "*", "[", "]", ":")
    For Each b In bad
        r = Replace(r, CStr(b), "")
    Next b
    r = Trim$(r)
    If Len(r) > 31 Then r = Left$(r, 31)
    If Len(r) = 0 Then r = "請求書"
    SafeSheetName = r
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not (w Is Nothing)
End Function

' 社名をシート名に。重複時は末尾にコードを付けて一意化(31文字以内)
Private Function UniqueSheetName(ByVal baseNm As String, ByVal code As String) As String
    Dim nm As String: nm = SafeSheetName(baseNm)
    If Not SheetExists(nm) Then UniqueSheetName = nm: Exit Function
    Dim sfx As String: sfx = "_" & code
    If Len(nm) + Len(sfx) > 31 Then nm = Left$(nm, 31 - Len(sfx))
    UniqueSheetName = nm & sfx
End Function

' 内訳の1行(ラベル + 金額を数式で)を書く
Private Sub PutSumF(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal fml As String)
    ws.Range("D" & r).Value = label
    ws.Range("D" & r).HorizontalAlignment = xlRight
    ws.Range("E" & r).Formula = fml
    ws.Range("E" & r).NumberFormat = "#,##0""円"""
End Sub

Private Sub WriteHeaderCell(ByVal rng As Range, ByVal text As String)
    rng.Value = text
    rng.Font.Bold = True
    rng.Font.Color = vbWhite
    rng.Interior.Color = ACCENT
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub HeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, cols))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = ACCENT
    End With
End Sub

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = nm
    End If
End Function

Private Function Sheet(ByVal nm As String) As Worksheet
    Set Sheet = ThisWorkbook.Worksheets(nm)
End Function

Private Sub RemoveGeneratedInvoices()
    Dim i As Long, ws As Worksheet
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If IsInvoiceSheet(ws) And ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
    Next i
    Application.DisplayAlerts = True
End Sub

Private Sub RemoveDefaultBlankSheets()
    Dim ws As Worksheet, nm As String
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        nm = ws.Name
        If nm <> S_START And nm <> S_CUST And nm <> S_SALES Then
            If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
                If ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
            End If
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub
