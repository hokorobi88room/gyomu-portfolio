' ============================================================
'  請求書デモ 自動組み立てスクリプト (Windows + Excel 用)
'  このファイルをダブルクリックすると、../src の .bas を取り込み、
'  サンプルデータ・ボタン付きの体験版ブックを組み立てて
'  「InvoiceForge_体験版.xlsm」を build フォルダに保存します。
'
'  事前に一度だけ Excel 側で許可が必要:
'   Excel > ファイル > オプション > トラスト センター
'    > トラスト センターの設定 > マクロの設定
'    >「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」に✓
'  (これは自分のPCで開発する人向けの設定。配布先には不要)
'
'  ※うまく動かない場合は同フォルダの「組み立て手順.txt」の
'    手動手順(3分)をご覧ください。
' ============================================================
Option Explicit

Dim fso, here, srcDir, outFile, xl, wb, comp, f, imported
Set fso = CreateObject("Scripting.FileSystemObject")
here   = fso.GetParentFolderName(WScript.ScriptFullName)
srcDir = fso.GetAbsolutePathName(here & "\..\src")
outFile = fso.BuildPath(here, "InvoiceForge_体験版.xlsm")

If Not fso.FolderExists(srcDir) Then
    MsgBox "src フォルダが見つかりません: " & srcDir, vbCritical, "組み立て失敗"
    WScript.Quit 1
End If

On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excel を起動できませんでした。Excel がインストールされたWindowsで実行してください。", vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add

' --- ../src の .bas をすべて取り込む ---
imported = 0
For Each f In fso.GetFolder(srcDir).Files
    If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
        On Error Resume Next
        wb.VBProject.VBComponents.Import f.Path
        If Err.Number <> 0 Then
            MsgBox "モジュール取り込みでエラー。" & vbCrLf & vbCrLf & _
                   "Excelのトラストセンターで" & vbCrLf & _
                   "「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」に" & vbCrLf & _
                   "チェックを入れてから、もう一度実行してください。" & vbCrLf & vbCrLf & _
                   "(詳細: " & Err.Description & ")", vbCritical, "組み立て失敗"
            wb.Close False : xl.Quit
            WScript.Quit 1
        End If
        On Error GoTo 0
        imported = imported + 1
    End If
Next

' --- 体験版ブックを組み立てる ---
On Error Resume Next
xl.Run "BuildDemoWorkbook"
If Err.Number <> 0 Then
    MsgBox "組み立てマクロの実行でエラー: " & Err.Description, vbCritical
    wb.Close False : xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- .xlsm として保存 (52 = xlOpenXMLWorkbookMacroEnabled) ---
If fso.FileExists(outFile) Then fso.DeleteFile outFile
wb.SaveAs outFile, 52
wb.Close False
xl.Quit

MsgBox "できました！" & vbCrLf & vbCrLf & _
       outFile & vbCrLf & vbCrLf & _
       "取り込んだモジュール: " & imported & " 個" & vbCrLf & vbCrLf & _
       "このファイルを開いて「スタート」シートの" & vbCrLf & _
       "緑ボタンを押すと動作を確認できます。", vbInformation, "組み立て完了"
