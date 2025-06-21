Dim fso, file

Set fso = CreateObject("Scripting.FileSystemObject")

' ファイルの保存先を指定（デスクトップに保存する例）
Set shell = CreateObject("WScript.Shell")
desktopPath = shell.SpecialFolders("Desktop")
Set file = fso.CreateTextFile(desktopPath & "\メモ帳でプログラミング\01\output.txt", True)

' テキストを書き込む
file.WriteLine "プログラムがやりました"

' ファイルを閉じる
file.Close

MsgBox "メモ帳に出力しました。", vbInformation, "完了"