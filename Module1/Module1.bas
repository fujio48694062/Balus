Attribute VB_Name = "Module1"
Option Explicit

Dim answer As String


Sub �o���X()

  MsgBox "���X�J�卲�u3���ԑ҂��Ă��B�v"
  MsgBox "���X�J�卲�u���Ԃ��B�����𕷂����B�v"
  
  answer = InputBox("���X�J�卲�u���Ԃ��B�����𕷂����B�v")
  
  If answer Like "�o���X*" Then '�u�o���X�v�̕�������܂ޏꍇ'
    MsgBox "���X�J�卲�u�ڂ��A�ڂ����`�I�v"
  Else
    MsgBox "���X�J�卲�u�o���X�v"
    MsgBox "���X�J�卲�u�ڂ��A�ڂ����`�I�v"
  End If

End Sub
