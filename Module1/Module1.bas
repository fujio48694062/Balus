Attribute VB_Name = "Module1"
Option Explicit

Dim answer As String


Sub �o���X()

  Worksheets("Sheet1").Select
  MsgBox "���X�J�卲�u3���ԑ҂��Ă��B�v"
  MsgBox "���X�J�卲�u���Ԃ��B�����𕷂����B�v"
  
  answer = InputBox("���X�J�卲�u���Ԃ��B�����𕷂����B�v")
  
  If answer Like "�o���X*" Then '�u�o���X�v�̕�������܂ޏꍇ'
    Worksheets("Sheet2").Select
    MsgBox "�p�Y�[&�V�[�^�u�o���X�v"
  Else
    Worksheets("Sheet3").Select
    MsgBox "���X�J�卲&�h�[���u�o���X�v"
  End If
  
  Worksheets("Sheet4").Select
  MsgBox "���X�J�卲�u�ڂ��A�ڂ����`�I�v"

End Sub
