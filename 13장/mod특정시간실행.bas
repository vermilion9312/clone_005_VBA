Attribute VB_Name = "modƯ���ð�����"
Option Explicit
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  1�ʸ��� �ѹ��� ����Ǿ� B2 �� ���� 1�� ������Ű�� ��ũ��
'   ���� ����: ������ ������ �ڵ������ �����ؾ���
'------------------------------------------------------------------------------------------
Dim setTime    '--// ������ �ð� ���

Sub sbTimer_Start()
Attribute sbTimer_Start.VB_Description = "1�ʸ��� �ѹ��� ����Ǿ� B2 �� ���� 1�� ������Ű�� ��ũ��"
Attribute sbTimer_Start.VB_ProcData.VB_Invoke_Func = " \n17"
    setTime = Now + TimeValue("00:00:01")         '--//�̺�Ʈ ������ �ð��� ���� �ð��� 1�ʸ� ���� ���
    Application.OnTime setTime, "sbTimer_Start"    '--//setTime �ð��� ��ũ�� ����
    
    Range("B2") = Range("B2").Value + 1      '--// ���� 1 �ʾ� ����
    Range("B2").NumberFormatLocal = "#,##0��"
End Sub


Sub sbTimer_Stop()
Attribute sbTimer_Stop.VB_Description = "sbTimer_Start �ڵ� ���� ����� ����"
Attribute sbTimer_Stop.VB_ProcData.VB_Invoke_Func = " \n17"
    On Error Resume Next
    '--// setTime �ð��� ������ sbTimer_Start ���ν��� ����
    Application.OnTime setTime, "sbTimer_Start", , False
End Sub

Sub sbTimer_Clear()
    Range("B2").Value = 0
    Call sbTimer_Start
End Sub






