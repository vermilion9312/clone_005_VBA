Attribute VB_Name = "mod���ڼ���������"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���ڿ����� ������ 1~4������ �����Ͽ�
'            ����(����:1), ������(����:2), �ѱ�(����:3), ��Ÿ����(����:4)
'            �� �����Ͽ� ��ȯ
'------------------------------------------------------------------------------------------
Function fn��������(���ڿ� As String, ���� As Integer) As String
Attribute fn��������.VB_Description = "���ڿ����� ������ 1~4������ �����Ͽ� \n����(����:1), ������(����:2), �ѱ�(����:3), ��Ÿ����(����:4)�� �����Ͽ� ��ȯ"
Attribute fn��������.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  Dim K As String
  '--// ����, ����, �ѱ�, ��Ÿ ���ڸ� ������ ����
  Dim NumStr As String, EngStr As String, HanStr As String, EtcStr As String  '��Ÿ ���ڵ��� �����
                                 
  For i = 1 To Len(���ڿ�)
      K = Mid(���ڿ�, i, 1)
      Select Case K
         Case "0" To "9"
           NumStr = NumStr & K
         Case "."
           NumStr = NumStr & K
         Case "A" To "Z"
           EngStr = EngStr & K
         Case "a" To "z"
           EngStr = EngStr & K
         Case "��" To "�P"    '�ѱ��� '��'�� ���� �۰� '�P'�� ���� ū ����
           HanStr = HanStr & K
         Case Else
           EtcStr = EtcStr & K
      End Select
  Next
  
  Select Case ����
      Case 1:          fn�������� = NumStr
      Case 2:          fn�������� = EngStr
      Case 3:          fn�������� = HanStr
      Case 4:          fn�������� = EtcStr
      Case Else:       fn�������� = "����"
  End Select
End Function

