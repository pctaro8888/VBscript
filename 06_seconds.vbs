Set shell = CreateObject("WScript.Shell")
timeoutSeconds = 5

shell.Popup(timeoutSeconds & "秒後にメッセージを出力します")

WScript.Sleep 1000 * timeoutSeconds
MsgBox "メッセージ", vbInformation, "出力確認"