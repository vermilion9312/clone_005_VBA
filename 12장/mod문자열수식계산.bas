Attribute VB_Name = "mod���ڿ����İ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���ڿ� ���� �� ���ڿ� ������ �̿��� ���ڸ� �����ϰ� ���
'------------------------------------------------------------------------------------------
Function fn���ڿ����(���ڿ��εȼ��� As String)
Attribute fn���ڿ����.VB_Description = "���ڿ� ���� �� ���ڿ� ������ �̿��� ���ڸ� �����ϰ� ���"
Attribute fn���ڿ����.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.Volatile
  
  Dim strResult As String, strTemp As String
  Dim i As Integer
  Const cOperator As String = "-+*/()^.%"
  
  For i = 1 To Len(���ڿ��εȼ���)
    strTemp = Mid(���ڿ��εȼ���, i, 1)
    If strTemp >= "0" And strTemp <= "9" Then
        strResult = strResult & strTemp
    ElseIf InStr(1, cOperator, strTemp) <> 0 Then
        strResult = strResult & strTemp
    End If
  Next
  
  fn���ڿ���� = Application.Evaluate(strResult)
End Function

'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���ڿ� ������ ������ ���
'            ���ڿ� ���뿡 ����, ������ �̿ܿ��� �Է��ϸ� 0���� ���
'------------------------------------------------------------------------------------------
Function fn�ܼ����ڿ����(���ڿ��εȼ��� As String)
Attribute fn�ܼ����ڿ����.VB_Description = "���ڿ� ������ ������ ��� ���ڿ� ���뿡 ����, ������ �̿ܿ��� �Է��ϸ� 0���� ���"
Attribute fn�ܼ����ڿ����.VB_ProcData.VB_Invoke_Func = " \n17"
  Application.Volatile
  
  �ܼ����ڿ���� = Application.Evaluate(���ڿ��εȼ���)
End Function


