Attribute VB_Name = "modMain"
'-----------------------------------------------------------------
'  ��   �� : ����� ���� ���� �����ϱ� ���� ��ũ��
'  �ۼ��� :  ��� ���� ��ũ�ο� VBA
'-----------------------------------------------------------------
Option Explicit

Sub �׷캰��Ʈ�и�()
Attribute �׷캰��Ʈ�и�.VB_ProcData.VB_Invoke_Func = " \n14"
   UserForm1.Show
End Sub

'-------------------------------------------------
'  ���� �޴��� ����� �߰��Ͽ� ����ϴ� ���
'  customUI.xml�� .rels ������ ������ �ʿ���
'  sbChooseMacro�� ���� �޴��� ��ư�� ����Ǿ� �����
'-------------------------------------------------
Sub sbChooseMacro(button As IRibbonControl)
  Select Case button.ID
    Case "customButton1"
      UserForm1.Show
    Case "customButton2"
      Call ��۾�ü���׷�з�
    Case "customButton3"
      Call �����������׷�з�
    Case Else
      MsgBox "�ش� ��ɰ� ����� �۾��� �����ϴ�.", vbInformation, "�ȳ�"
  End Select
End Sub

