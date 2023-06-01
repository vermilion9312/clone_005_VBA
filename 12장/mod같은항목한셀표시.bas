Attribute VB_Name = "mod같은항목한셀표시"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  수직 열 방향으로 두 영역의 자료를 입력받아 지정된 항목과
'            같은 자료의 값을 콤마로 분리하여 반환
'------------------------------------------------------------------------------------------

Option Compare Text

Function fn한셀표시(찾을값, 검색범위 As Range, 가져올자료범위 As Range) As String
Attribute fn한셀표시.VB_Description = "수직 열 방향으로 두 영역의 자료를 입력받아 지정된 항목과 같은 자료의 값을 콤마로 분리하여 반환"
Attribute fn한셀표시.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim strOut As String
  Dim i As Long
  
  Dim OutData As New Collection
  Dim varK As Variant
   
  If 검색범위.Rows.Count <> 가져올자료범위.Rows.Count Then
    fn한셀표시 = "검색범위 행과 가져올자료 행의 수가 동일해야 합니다."
    Exit Function
  End If
    
  On Error Resume Next
  
  For i = 1 To 검색범위.Rows.Count
    If 검색범위.Item(i) = 찾을값 Then
        OutData.Add Item:=가져올자료범위.Item(i), key:=CStr(가져올자료범위.Item(i))
    End If
  Next

  For Each varK In OutData
    If strOut = "" Then
         strOut = varK
    Else
         strOut = strOut & ", " & varK
    End If
  Next
  
  fn한셀표시 = strOut
End Function

