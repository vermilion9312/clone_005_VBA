Attribute VB_Name = "mod����Ű�����"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� : Application.OnKey �޼��带 �̿��Ͽ� Ư�� Ű ���տ� ����� �����ϱ�
'   �������� : �ش� ����� �ڵ����� ����ǵ��� �� ����
'                 ThisWorkbook(����_����_����) ���� 'Workbook_Open' �̺�Ʈ�� �̿���
'                 ����ǵ��� �ؾ� ��
'------------------------------------------------------------------------------------------
Option Explicit

Sub sbŰ��ɼ���()
Attribute sbŰ��ɼ���.VB_Description = "Application.OnKey �޼��带 �̿��Ͽ� Ư�� Ű ���տ� ����� �����ϱ�"
Attribute sbŰ��ɼ���.VB_ProcData.VB_Invoke_Func = " \n17"
    Application.OnKey "+{F9}", "sb�޽���"
    MsgBox "<Shift>+<F9> Ű�� ���� ���Ű ������ �����Ǿ����ϴ�."
End Sub

Sub sbŰ��ɻ���()
Attribute sbŰ��ɻ���.VB_Description = "Application.OnKey �޼��带 �̿��Ͽ� Ư�� Ű ���տ� ����� �����ϱ�"
Attribute sbŰ��ɻ���.VB_ProcData.VB_Invoke_Func = " \n17"
    Application.OnKey "+{F9}"
    MsgBox "<Shift>+<F9> Ű�� ���� ���Ű ������ �����Ǿ����ϴ�."
End Sub

Sub sb�޽���()
   MsgBox Time, vbInformation, "���� �ð�"
End Sub

