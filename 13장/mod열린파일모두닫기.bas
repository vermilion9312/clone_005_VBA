Attribute VB_Name = "mod�������ϸ�δݱ�"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���� ������ ������ ������ �����ִ� ��� �ش� ������ ���� ���θ� Ȯ���� �� �ݴ� ó��
'------------------------------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------------------------------
'  ���� ������ ������ ������ �����ִ� ��� �ش� ������ ���� ���θ� Ȯ���� �� �ݴ� ó��
'  ó������ �޽����� ����� �� �� ������ Ȯ�� ���� �����ϰ� ������ �� ���� �Ű����� ó��
'--------------------------------------------------------------------------------------
Sub sbClose()
Attribute sbClose.VB_Description = "���� ������ ������ ������ �����ִ� ��� �ش� ������ ���� ���θ� Ȯ���� �� �ݴ� ó��"
Attribute sbClose.VB_ProcData.VB_Invoke_Func = " \n17"
    Dim i As Integer
    Dim K As Workbook
    Dim strMsg As String, bSave As Boolean
    
   If Workbooks.Count = 1 Then Exit Sub
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
On Error GoTo End_Rtn
   i = MsgBox("���� ������ ������ ������ ������ ����" & Workbooks.Count - 1 & "�����Դϴ�." & vbCrLf & _
       "�� �������� �ݰ� �۾��ؾ� �մϴ�. �����ϰ� �������?" & vbCrLf & _
       "������-�����ϰ� �ݱ�" & vbCr & "���ƴϿ���-���������ʰ� �ݱ�" & vbCr & "����ҡ�-�۾����", vbQuestion + vbYesNoCancel, "���� �ݱ� Ȯ��")
   
   Select Case i
       Case vbYes
           bSave = True: strMsg = "���� ������ ������ �����ϰ� �ݴ� ���Դϴ�. "
       Case vbNo
           bSave = False: strMsg = "���� ������ ������ �����ϰ� �ݴ� ���Դϴ�. "
       Case Else
         End
   End Select
    
   For Each K In Workbooks
        If K.Name <> ThisWorkbook.Name Then
            If K.ReadOnly Then   '--// �б� �������� ���� ������ üũ
                K.Close False
            Else
                K.Close bSave
            End If
        End If
   Next
    
End_Rtn:
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   If Err.Number = 0 Then
      MsgBox "�۾��� ���������� �Ϸ��߽��ϴ�.", vbInformation, "�۾��Ϸ�"
   Else
      MsgBox "�۾� �� ������ ���� ������ �߻��߽��ϴ�." & vbCrLf & _
               Err.Description, vbCritical, "����"
   End If
End Sub


