Attribute VB_Name = "mod���α׷�����"
Option Explicit

Sub sb�������()
   Dim Cal
   On Error Resume Next
   Cal = Shell("calc.exe", vbNormalFocus)
   If Cal = "" Then MsgBox "������� �ʾҽ��ϴ�."
End Sub
