Attribute VB_Name = "Module1"
Option Explicit

Dim answer As String


Sub バルス()

  Worksheets("Sheet1").Select
  MsgBox "ムスカ大佐「3分間待ってやる。」"
  MsgBox "ムスカ大佐「時間だ。答えを聞こう。」"
  
  answer = InputBox("ムスカ大佐「時間だ。答えを聞こう。」")
  
  If answer Like "バルス*" Then '「バルス」の文字列を含む場合'
    Worksheets("Sheet2").Select
    MsgBox "パズー&シータ「バルス」"
  Else
    Worksheets("Sheet3").Select
    MsgBox "ムスカ大佐&ドーラ「バルス」"
  End If
  
  Worksheets("Sheet4").Select
  MsgBox "ムスカ大佐「目が、目がぁ～！」"

End Sub
