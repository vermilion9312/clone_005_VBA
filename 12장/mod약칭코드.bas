Attribute VB_Name = "mod��Ī�ڵ�"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� : �ܾ� ù���ڷ� ��� �����
'------------------------------------------------------------------------------------------
Option Explicit

Function fn��Ī�ڵ�(���ڿ� As Range)
Attribute fn��Ī�ڵ�.VB_Description = "�ܾ� ù���ڷ� ��� �����"
Attribute fn��Ī�ڵ�.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim str�ڵ� As String, strTemp As String
   Dim i As Integer
   
   strTemp = Trim(���ڿ�)
   
   Do
      i = i + 1
      str�ڵ� = str�ڵ� & Mid(strTemp, i, 1)
      i = InStr(i, strTemp, " ")
   Loop Until i = 0
   
   fn��Ī�ڵ� = UCase(str�ڵ�)
End Function

