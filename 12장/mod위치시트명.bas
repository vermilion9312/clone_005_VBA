Attribute VB_Name = "mod��ġ��Ʈ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ��ȣ�� Ư����Ʈ�̸� ��ȯ
'------------------------------------------------------------------------------------------
Function fn��ġ��Ʈ��(��Ʈ��ȣ As Integer)
Attribute fn��ġ��Ʈ��.VB_Description = "��ȣ�� Ư����Ʈ�̸� ��ȯ"
Attribute fn��ġ��Ʈ��.VB_ProcData.VB_Invoke_Func = " \n14"
   Application.Volatile
   
   If ��Ʈ��ȣ > Application.ThisCell.Parent.Parent.Sheets.Count Then
      fn��ġ��Ʈ�� = "����"
   Else
      fn��ġ��Ʈ�� = Application.ThisCell.Parent.Parent.Sheets(��Ʈ��ȣ).Name
   End If
End Function
