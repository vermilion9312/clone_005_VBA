Attribute VB_Name = "mod���������"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ��� ������ �۲� ���� ������ ���� ������ �ָ�
'            ��� �������� Ư�� ���� �� ���� �հ踸 ���
'------------------------------------------------------------------------------------------

Function fn���������(������ As Range, ������ As Range, Optional �Լ��̸� As String = "SUM")
Attribute fn���������.VB_Description = "��� ������ �۲� ���� ������ ��, ����� �Լ��̸�\n(��; SUM, AVERAGE, COUNT, COUNTA, MAX, MIN ��)��\n������ �ָ� ���� �ȿ��� ������ �۲� ���� ���� ���� ã��\n�ش� �Լ��� ����� ����� ��ȯ�մϴ�."
Attribute fn���������.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim K As Range, rngCal As Range
   Dim Result As Double
   
   For Each K In ������
      If K.Font.Color = ������.Font.Color Then
         If rngCal Is Nothing Then
            Set rngCal = K
         Else
            Set rngCal = Union(rngCal, K)
         End If
      End If
   Next
   If Not rngCal Is Nothing Then
      Result = Application.Evaluate(�Լ��̸� & "(" & rngCal.Address & ")")
   End If
   fn��������� = Result
End Function

