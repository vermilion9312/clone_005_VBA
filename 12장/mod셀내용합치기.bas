Attribute VB_Name = "mod��������ġ��"
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� : Ư�� ������ ������ �ߺ������Ͽ� ������������ �Ѽ��� ǥ��
'------------------------------------------------------------------------------------------
Option Explicit

Function fn��������ġ��(���� As Range) As String
Attribute fn��������ġ��.VB_Description = "������ �� ���� ������ �� ����"
Attribute fn��������ġ��.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim OutData As New Collection
  Dim varK
  Dim strOut As String
  Dim i As Long, k As Long
  
  On Error Resume Next
  For i = 1 To ����.Cells.Count
    varK = ����.Cells(i)
    If varK <> "" Then OutData.Add Item:=varK, Key:=CStr(varK)
  Next
  On Error GoTo 0
  
  For i = 1 To OutData.Count - 1
      For k = i + 1 To OutData.Count
        If OutData(i) > OutData(k) Then
          varK = OutData(k)
          OutData.Remove k
          OutData.Add Item:=varK, Key:=CStr(varK), before:=i
        End If
      Next
  Next
  
  strOut = ""
  For Each varK In OutData
    strOut = strOut & IIf(strOut = "", "", ",") & varK
  Next
  fn��������ġ�� = strOut
End Function
