Dim answer
answer = MsgBox("プログラミングはお好きですか？", vbYesNo, "確認")

If response = vbYes Then
    MsgBox "保存します"
Else
    MsgBox "キャンセルしました"
End If