Attribute VB_Name = "mod�����׸��Ѽ�ǥ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���� �� �������� �� ������ �ڷḦ �Է¹޾� ������ �׸��
'            ���� �ڷ��� ���� �޸��� �и��Ͽ� ��ȯ
'------------------------------------------------------------------------------------------

Option Compare Text

Function fn�Ѽ�ǥ��(ã����, �˻����� As Range, �������ڷ���� As Range) As String
Attribute fn�Ѽ�ǥ��.VB_Description = "���� �� �������� �� ������ �ڷḦ �Է¹޾� ������ �׸�� ���� �ڷ��� ���� �޸��� �и��Ͽ� ��ȯ"
Attribute fn�Ѽ�ǥ��.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim strOut As String
  Dim i As Long
  
  Dim OutData As New Collection
  Dim varK As Variant
   
  If �˻�����.Rows.Count <> �������ڷ����.Rows.Count Then
    fn�Ѽ�ǥ�� = "�˻����� ��� �������ڷ� ���� ���� �����ؾ� �մϴ�."
    Exit Function
  End If
    
  On Error Resume Next
  
  For i = 1 To �˻�����.Rows.Count
    If �˻�����.Item(i) = ã���� Then
        OutData.Add Item:=�������ڷ����.Item(i), key:=CStr(�������ڷ����.Item(i))
    End If
  Next

  For Each varK In OutData
    If strOut = "" Then
         strOut = varK
    Else
         strOut = strOut & ", " & varK
    End If
  Next
  
  fn�Ѽ�ǥ�� = strOut
End Function

