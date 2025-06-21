'-----------------------------------------------------------
' 機能名　：画面ロック確認＋カウントダウン＋キャンセル機能付き
' 機能概要：〇秒後にPCをロック、途中でキャンセル可能
' 作成日　：2025.06.21
'-----------------------------------------------------------
Set shell = CreateObject("WScript.Shell")
timeoutSeconds = 5

response = shell.Popup(timeoutSeconds & "秒後に画面をロックします。", timeoutSeconds, "画面ロック", vbYesNo + 48)

If response = vbNo Then
    MsgBox "画面ロックはキャンセルされました。", vbInformation
Else
    settime = 1000 * timeoutSeconds
	WScript.Sleep settime
    shell.Run "rundll32.exe user32.dll,LockWorkStation", 0, False
End If