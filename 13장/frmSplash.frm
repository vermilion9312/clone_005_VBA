VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "Excel ��ũ�ο� VBA �����������ϱ�"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dtTime As Date      '--// ���� ������ �ð��� ����

Private Sub UserForm_Initialize()
'    '--// ���� ������ �� 5���Ŀ� ��������� ��
   dtTime = Now() + TimeValue("0:0:5")
    Application.OnTime dtTime, "sb���ݱ�"
End Sub

'--// ���� ���� �� ���ؿ����Ǿ��ִ� �۾��� ���
Private Sub UserForm_Terminate()
   If Now < dtTime Then
      Application.OnTime dtTime, "sb���ݱ�", Schedule:=False
   End If
End Sub

