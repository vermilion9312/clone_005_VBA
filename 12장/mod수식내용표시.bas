Attribute VB_Name = "mod���ĳ���ǥ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ������ ���� ���� ������ �ؽ�Ʈ�� ǥ��
'------------------------------------------------------------------------------------------
Function fn���ĺ���(�� As Range)
Attribute fn���ĺ���.VB_Description = "������ ���� ���� ������ �ؽ�Ʈ�� ǥ��"
Attribute fn���ĺ���.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim strTemp As String

    If ��.HasArray Then
        strTemp = "{" & ��.Formula & "}"
    ElseIf ��.HasFormula Then
        strTemp = ��.Formula
    Else
        strTemp = ""
    End If
    fn���ĺ��� = strTemp
End Function


