Attribute VB_Name = "mod���ͳݿ���"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� /  �̵���(bofb@naver.com) ��
'   ��� :  ���ͳ� ���� ���� Ȯ��
'            URL �ּҸ� �Է��Ͽ� �ش� ����Ʈ ���� ���� ���� �Ǵ�
'            URL �ּ� ������ ���ͳ� ���� ����/�Ұ��� üũ
'------------------------------------------------------------------------------------------

Option Explicit

Private Const FLAG_ICC_FORCE_CONNECTION = &H1
#If VBA7 And Win64 Then      '--// 64bit
      Private Declare PtrSafe Function InternetCheckConnection Lib "wininet.dll" _
            Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, _
            ByVal dwflags As Long, ByVal dwReserved As Long) As Long
      Private Declare PtrSafe Function InternetAttemptConnect Lib "wininet" _
            (ByVal dwReserved As Long) As Long

#Else
      Private Declare Function InternetCheckConnection Lib "wininet.dll" _
            Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, _
            ByVal dwflags As Long, ByVal dwReserved As Long) As Long
      Private Declare Function InternetAttemptConnect Lib "wininet" _
            (ByVal dwReserved As Long) As Long

#End If

Function fn���ͳݻ���(Optional URL As String)
Attribute fn���ͳݻ���.VB_Description = "���ͳ� ���� ���� Ȯ��\nURL �ּҸ� �Է��Ͽ� �ش� ����Ʈ ���� ���� ���θ� [���󿬰�]/[���ӺҰ�]�� ǥ��\nURL �ּ� ������ ���ͳ� ���� ���� ���θ� [���ᰡ��]/[����Ұ�]�� ǥ��"
Attribute fn���ͳݻ���.VB_ProcData.VB_Invoke_Func = " \n14"
   If InternetAttemptConnect(0) = 0 Then
      If URL = "" Then
         fn���ͳݻ��� = "���ᰡ��"
      ElseIf InternetCheckConnection(URL, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
         fn���ͳݻ��� = "���ӺҰ�"
      Else
         fn���ͳݻ��� = "���󿬰�"
      End If
   Else
       fn���ͳݻ��� = "����Ұ�"
   End If
End Function



