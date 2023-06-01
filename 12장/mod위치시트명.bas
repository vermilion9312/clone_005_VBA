Attribute VB_Name = "mod위치시트명"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  번호로 특정시트이름 반환
'------------------------------------------------------------------------------------------
Function fn위치시트명(시트번호 As Integer)
Attribute fn위치시트명.VB_Description = "번호로 특정시트이름 반환"
Attribute fn위치시트명.VB_ProcData.VB_Invoke_Func = " \n14"
   Application.Volatile
   
   If 시트번호 > Application.ThisCell.Parent.Parent.Sheets.Count Then
      fn위치시트명 = "없음"
   Else
      fn위치시트명 = Application.ThisCell.Parent.Parent.Sheets(시트번호).Name
   End If
End Function
