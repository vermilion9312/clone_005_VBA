Attribute VB_Name = "mod내용유지셀병합"
Option Base 1
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : 선택한 범위를 셀 병합하고 범위의 각 셀 내용을 하나의 텍스트로 병합하여 표시
'------------------------------------------------------------------------------------------

Sub sb내용유지셀병합()
Attribute sb내용유지셀병합.VB_Description = "선택한 범위를 셀 병합하고 범위의 각 셀 내용을 하나의 텍스트로 병합하여 표시"
Attribute sb내용유지셀병합.VB_ProcData.VB_Invoke_Func = " \n17"
  Dim rngData 'As Range
  Dim i As Integer, j As Integer
  Dim xCell() As String
  Dim strResult As String
  
  Application.DisplayAlerts = False
On Error GoTo End_Rtn
  
  If Selection.Areas.Count <> 1 Then
      MsgBox "2개 이상의 범위 영역이 지정되었습니다." & vbCr & _
                    "범위를 다시 선택해 주세요!"
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
  
  iconf = MsgBox("셀 병합을 되돌릴까요?", vbYesNo, "셀 병합 해제")
  
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
  If Err > 0 Then MsgBox Err.Description, , "작업오류"
  Application.DisplayAlerts = False
End Sub


