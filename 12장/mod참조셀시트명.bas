Attribute VB_Name = "mod��������Ʈ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���� ���� ��Ʈ �Ǵ� ���ϸ��� ��ȯ
'------------------------------------------------------------------------------------------
Function fn����������(Optional �� As Range, Optional IsSheetName As Boolean = True)
Attribute fn����������.VB_Description = "���� ���� ��Ʈ �Ǵ� ���ϸ��� ��ȯ"
Attribute fn����������.VB_ProcData.VB_Invoke_Func = "\n14"
   Application.Volatile
   Dim sh As Worksheet
   
   If �� Is Nothing Then
      Set sh = Application.ThisCell.Parent
   Else
      Set sh = ��.Parent
   End If
   fn���������� = IIf(IsSheetName, sh.Name, sh.Parent.Name)
End Function
