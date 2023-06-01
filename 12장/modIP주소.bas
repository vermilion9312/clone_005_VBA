Attribute VB_Name = "modIP�ּ�"

'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :   IP �ּҸ� ��� �Լ� : fnGetIPAddress
'------------------------------------------------------------------------------------------
'--// Microsoft WMI Scripting Library  �� �����Ͽ� ó���ϴ� ���
Function fnGetIPAddress()
Attribute fnGetIPAddress.VB_Description = "���� ��ǻ���� IP �ּҸ� IPv4, IPv6 ������� ��ȯ"
Attribute fnGetIPAddress.VB_ProcData.VB_Invoke_Func = "\n14"
   ' ��ǻ�� ������ ���� ��ǻ��(local computer)�� �ƴ�
    Const strComputer As String = "."
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' WMI service�� ����
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' TCP/IP�� ������ ����͵��� ������ ������
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' TCP/IP�� ������ ����͵鿡 �ο��� IP �ּҸ� ������
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
        End If
    Next

    fnGetIPAddress = strIPAddress
End Function


