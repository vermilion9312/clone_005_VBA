Attribute VB_Name = "mod_�����޴��߰�"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���� �޴��� �߰��� �� �ش� ����� Ŭ���� �� ������
'            ���ν����� �ۼ�
'------------------------------------------------------------------------------------------
Option Explicit

Sub RibbonControl_Click(button As Office.IRibbonControl)
   Select Case button.ID
      Case "Button1": Call sbMsg_Time
      Case "Button2": Call sbMsg_Date
      
      Case Else: Call btnMsg(button.ID)
      
   End Select
End Sub

Sub sbMsg_Time()
   MsgBox "���� �ð� :" & Time, vbInformation, "���� �޴� ����"
End Sub

Sub sbMsg_Date()
   MsgBox "���� ��¥ :" & Date, vbInformation, "���� �޴� ����"
End Sub

Sub btnMsg(btnId As String)
   MsgBox btnId & "�� ���� ó���� �������� �ʾҽ��ϴ�.", vbCritical
End Sub
