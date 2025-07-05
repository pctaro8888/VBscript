'==============================
'名称：自動スクリーンショット機能
'機能：URLを読み取って、GUIなし（ヘッドレス）でスクリーンショットを取得する
'備考：Edge/Chromeのヘッドレスモードを使用
'処理： ①変数・定数定義
'       ②フォルダの存在チェック
'       ③URL読み込み
'       ④開始メッセージ
'       ⑤ループ処理
'           5-1.定数・変数定義
'           5-2.ヘッドレスモードのオプション設定
'           5-3.実行
'           5-4.エラーハンドリング
'       ⑥終了メッセージ
'==============================

'===== ① 変数・定数定義 =====
Dim shell, fso, desktopPath, saveDir
Const EDGE_PATH = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe""" ' edgeの格納先

Set shell = CreateObject("WScript.Shell") ' コマンド実行用
Set fso   = CreateObject("Scripting.FileSystemObject") ' ファイル操作用

Const SAVE_SUBDIR = "\メモ帳でプログラミング\20250705\shots" ' 保存先サブパス
desktopPath = shell.SpecialFolders("Desktop") ' ユーザーのデスクトップパス
saveDir     = desktopPath & SAVE_SUBDIR       ' フルパス組み立て

'===== ② フォルダ存在チェック・作成 =====
If Not fso.FolderExists(saveDir) Then fso.CreateFolder(saveDir)

'===== ③ URL 読み込み =====
Dim urls, url

urls = Array( _
    "https://radiko.jp/", _
    "https://www.ne.jp/asahi/beat/non/mori/", _
    "https://dailyportalz.jp/" _
)


'===== ④ 開始メッセージ =====
MsgBox UBound(urls) + 1 & " 件のスクリーンショットを取得します。"

'===== ⑤ メイン処理ループ =====
Dim i, failName, outputFile, cmd, exitCode, totalOk, totalNg
totalOk = 0 : totalNg = 0

For i = 0 To UBound(urls)
    url = urls(i)
    MsgBox url ' 実行前に確認表示（省略可）
    failName = i

    outputFile = """" & saveDir & "\" & failName & ".png" & """"

    '===== ⑤-2 ヘッドレスモードのオプション =====
    cmd = EDGE_PATH _
        & " --headless=new" _
        & " --window-size=1280,2000" _
        & " --run-all-compositor-stages-before-draw" _
        & " --virtual-time-budget=10000" _
        & " --screenshot=" & outputFile _
        & " " & url

    '===== ⑤-3 実行とハンドリング =====
    exitCode = shell.Run(cmd, 1, True) ' 第2引数：1=通常表示（非0だと黒い画面一瞬表示）

    If exitCode = 0 Then
        totalOk = totalOk + 1
    Else
        totalNg = totalNg + 1
        MsgBox "エラー (" & exitCode & ") : " & url
    End If
Next

'===== ⑥ 完了メッセージ =====
MsgBox "完了！ 成功: " & totalOk & " 件 / 失敗: " & totalNg & " 件"
