Attribute VB_Name = "mod��������������"
Option Base 1
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� : ������ ������ �� �����ϰ� ������ �� �� ������ �ϳ��� �ؽ�Ʈ�� �����Ͽ� ǥ��
'------------------------------------------------------------------------------------------

Sub sb��������������()
Attribute sb��������������.VB_Description = "������ ������ �� �����ϰ� ������ �� �� ������ �ϳ��� �ؽ�Ʈ�� �����Ͽ� ǥ��"
Attribute sb��������������.VB_ProcData.VB_Invoke_Func = " \n17"
  Dim rngData 'As Range
  Dim i As Integer, j As Integer
  Dim xCell() As String
  Dim strResult As String
  
  Application.DisplayAlerts = False
On Error GoTo End_Rtn
  
  If Selection.Areas.Count <> 1 Then
      MsgBox "2�� �̻��� ���� ������ �����Ǿ����ϴ�." & vbCr & _
                    "������ �ٽ� ������ �ּ���!"
      Exit Sub
  End If
  
  ReDim xCell(Selection.Rows.Count, Selection.Columns.Count)
  For i = 1 To UBound(xCell, 1)
    For j = 1 To UBound(xCell, 2)
        xCell(i, j) = Selection.Cells(i, j)
    Next j
  Next i
  
  For i = 1 To UBound(xCell, 1)
    If i > 1 Then strResult = strResult & vbLf
    For j = 1 To UBound(xCell, 2)
        strResult = strResult & xCell(i, j)
    Next j
  Next i
  
  With Selection
    .MergeCells = False
    .Merge
  End With
  
  Selection.Cells(1, 1) = strResult
  Application.Goto Selection
  
  iconf = MsgBox("�� ������ �ǵ������?", vbYesNo, "�� ���� ����")
  
  If iconf = vbYes Then
    With Selection
      .MergeCells = True
      .UnMerge
    End With
    For i = 1 To UBound(xCell, 1)
      For j = 1 To UBound(xCell, 2)
          Selection.Cells(i, j) = xCell(i, j)
      Next j
    Next i
  End If
  
End_Rtn:
  If Err > 0 Then MsgBox Err.Description, , "�۾�����"
  Application.DisplayAlerts = False
End Sub


