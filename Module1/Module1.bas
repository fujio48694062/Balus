Attribute VB_Name = "Module1"
Option Explicit

Dim answer As String


Sub バルス()

  MsgBox "ムスカ大佐「3分間待ってやる。」"
  MsgBox "ムスカ大佐「時間だ。答えを聞こう。」"
  
  answer = InputBox("ムスカ大佐「時間だ。答えを聞こう。」")
  
  If answer Like "バルス*" Then '「バルス」の文字列を含む場合'
    MsgBox "ムスカ大佐「目が、目がぁ～！」"
  Else
    MsgBox "ムスカ大佐「バルス」"
    MsgBox "ムスカ大佐「目が、目がぁ～！」"
  End If

End Sub
