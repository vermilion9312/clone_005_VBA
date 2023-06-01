VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "Excel 매크로와 VBA 무작정따라하기"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dtTime As Date      '--// 폼이 닫히는 시간을 저장

Private Sub UserForm_Initialize()
'    '--// 폼을 실행한 후 5초후에 사라지도록 함
   dtTime = Now() + TimeValue("0:0:5")
    Application.OnTime dtTime, "sb폼닫기"
End Sub

'--// 폼이 닫힐 때 실해예정되어있던 작업을 취소
Private Sub UserForm_Terminate()
   If Now < dtTime Then
      Application.OnTime dtTime, "sb폼닫기", Schedule:=False
   End If
End Sub

