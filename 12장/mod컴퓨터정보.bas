Attribute VB_Name = "mod��ǻ������"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� : ������� ��ǻ���� ��ǻ�͸�� �α��� ����� ID�� ǥ��
'------------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------------------
'   �α����� ����� ID�� ��ǻ�� �̸�, ���� ���� ���� API �Լ�
'   ���� 2010 �̻󿡼� ���� ������ 32bit�� ���� 64bit�� �� �ٸ��� ó��
'------------------------------------------------------------------------------------------
'  VBA7�� 2010 �̻��� VBA�� �ǹ̷� ���� ������ VBA �ڵ������� ���� �� ���
'  Win64�� 32bit���� 64bit ���� ������ �� ���
#If VBA7 And Win64 Then      '--// 64bit
   Private Declare PtrSafe Function GetComputerName Lib "kernel32" _
            Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
   Private Declare PtrSafe Function GetUserName Lib "advapi32" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
   Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

#Else                '--//32Bit
   Private Declare Function GetComputerName Lib "kernel32" _
            Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function GetUserName Lib "advapi32" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
#End If

'------------------------------------------------------------------------------------------
'    ��ǻ���̸��� ��ȯ
'------------------------------------------------------------------------------------------
Function fn��ǻ�͸�()
Attribute fn��ǻ�͸�.VB_Description = "��ǻ���̸��� ��ȯ"
Attribute fn��ǻ�͸�.VB_ProcData.VB_Invoke_Func = " \n14"
      Dim tmp As String
      Application.Volatile
     
      tmp = Space$(256)
      
   GetComputerName tmp, 256
   fn��ǻ�͸� = Left$(tmp, lstrlenW(StrPtr(tmp)))
End Function

'------------------------------------------------------------------------------------------
' ��ǻ�� �α��� ����� ID ��ȯ
'------------------------------------------------------------------------------------------
Function fn��ǻ�ͻ����()
Attribute fn��ǻ�ͻ����.VB_Description = "��ǻ�� �α��� ����� ID ��ȯ"
Attribute fn��ǻ�ͻ����.VB_ProcData.VB_Invoke_Func = " \n14"
      Dim tmp As String
      Application.Volatile
      
      tmp = Space$(256)
      
   GetUserName tmp, 256
   fn��ǻ�ͻ���� = Left$(tmp, lstrlenW(StrPtr(tmp)))
End Function



