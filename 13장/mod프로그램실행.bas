Attribute VB_Name = "mod프로그램실행"
Option Explicit

Sub sb계산기실행()
   Dim Cal
   On Error Resume Next
   Cal = Shell("calc.exe", vbNormalFocus)
   If Cal = "" Then MsgBox "실행되지 않았습니다."
End Sub
