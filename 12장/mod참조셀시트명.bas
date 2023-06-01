Attribute VB_Name = "mod참조셀시트명"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  셀이 속한 시트 또는 파일명을 반환
'------------------------------------------------------------------------------------------
Function fn참조셀정보(Optional 셀 As Range, Optional IsSheetName As Boolean = True)
Attribute fn참조셀정보.VB_Description = "셀이 속한 시트 또는 파일명을 반환"
Attribute fn참조셀정보.VB_ProcData.VB_Invoke_Func = "\n14"
   Application.Volatile
   Dim sh As Worksheet
   
   If 셀 Is Nothing Then
      Set sh = Application.ThisCell.Parent
   Else
      Set sh = 셀.Parent
   End If
   fn참조셀정보 = IIf(IsSheetName, sh.Name, sh.Parent.Name)
End Function
