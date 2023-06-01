Attribute VB_Name = "mod내용유지셀병합"
Option Base 1
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : 선택한 범위를 셀 병합하고 범위의 각 셀 내용을 하나의 텍스트로 병합하여 표시
'------------------------------------------------------------------------------------------

Sub sb내용유지셀병합()
Attribute sb내용유지셀병합.VB_Description = "선택한 범위를 셀 병합하고 범위의 각 셀 내용을 하나의 텍스트로 병합하여 표시"
Attribute sb내용유지셀병합.VB_ProcData.VB_Invoke_Func = " \n17"
  Dim i As Integer, j As Integer
  Dim xCell() As String     '--// 선택한 영역의 셀 개수에 따라 크기 지정
  Dim strResult As String
  
  '--// 엑셀의 경고창을 표시하지 않도록 설정
  Application.DisplayAlerts = False
On Error GoTo End_Rtn
  
  '--// 선택 영역이 비 연속적인 영역인 경우 작업을 중단
  If Selection.Areas.Count <> 1 Then
      MsgBox "2개 이상의 범위 영역이 지정되었습니다." & vbCr & _
                    "범위를 다시 선택해 주세요!"
      Exit Sub
  End If
  
  '--//  선택한 영역의 행과 열 개수에 따라 xCell 배열 변수 크기를 지정
  '      병합 취소를 대비하여 셀의 내용을 xCell 에 저장
  '  strResult 변수에는 영역안의 내용을 취합하기 위해, 행이 변경되면 줄변경하고 열이 변경되면 공백을 입력하여 내용 연결
  ReDim xCell(Selection.Rows.Count, Selection.Columns.Count)
  For i = 1 To UBound(xCell, 1)
    If i > 1 Then strResult = strResult & vbLf
    For j = 1 To UBound(xCell, 2)
        If j > 1 Then strResult = strResult & " "
        strResult = strResult & Selection.Cells(i, j)
        
        xCell(i, j) = Selection.Cells(i, j)
    Next j
  Next i
  
  '--// 셀 병합 처리 (Merge 메서드 대신 Selection.MergeCells = True 로 속성을 사용해도 됨)
  Selection.Merge
  
  Selection.Cells(1, 1) = strResult
  Application.Goto Selection
  
  iconf = MsgBox("셀 병합을 되돌릴까요?", vbYesNo, "셀 병합 해제")
  
  If iconf = vbYes Then
    '--// 셀 병합 해제 (UnMerge 메서드 대신 Selection.MergeCells = False 로 속성을 사용해도 됨)
    Selection.UnMerge
    
    '--// 병합을 해제한 후 보관해 두었던 셀 내용을 선택 영역에 저장
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


