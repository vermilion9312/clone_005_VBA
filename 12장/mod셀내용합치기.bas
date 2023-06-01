Attribute VB_Name = "mod셀내용합치기"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : 특정 영역의 내용을 중복제거하여 오름차순으로 한셀에 표시
'------------------------------------------------------------------------------------------
Option Explicit

Function fn셀내용합치기(범위 As Range) As String
Attribute fn셀내용합치기.VB_Description = "범위내 셀 내용 병합한 후 정렬"
Attribute fn셀내용합치기.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim OutData As New Collection
  Dim varK
  Dim strOut As String
  Dim i As Long, k As Long
  
  On Error Resume Next
  For i = 1 To 범위.Cells.Count
    varK = 범위.Cells(i)
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
  fn셀내용합치기 = strOut
End Function
