Attribute VB_Name = "modIP주소"

'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :   IP 주소를 얻는 함수 : fnGetIPAddress
'------------------------------------------------------------------------------------------
'--// Microsoft WMI Scripting Library  를 참조하여 처리하는 방법
Function fnGetIPAddress()
Attribute fnGetIPAddress.VB_Description = "현재 컴퓨터의 IP 주소를 IPv4, IPv6 방식으로 반환"
Attribute fnGetIPAddress.VB_ProcData.VB_Invoke_Func = "\n14"
   ' 컴퓨터 명으로 로컬 컴퓨터(local computer)는 아님
    Const strComputer As String = "."
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' WMI service에 연결
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' TCP/IP가 가능한 어댑터들의 정보를 가져옴
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' TCP/IP가 가능한 어댑터들에 부여된 IP 주소를 가져옴
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
        End If
    Next

    fnGetIPAddress = strIPAddress
End Function


