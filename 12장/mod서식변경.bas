Attribute VB_Name = "mod서식변경"
Option Explicit
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  셀안의 특정 단어만 찾아 서식을 지정할 때 사용
'   함께 사용해야할 자료 : modPictureSaveAPI , frm서식
'------------------------------------------------------------------------------------------

Sub sb특정단어서식변경()
Attribute sb특정단어서식변경.VB_Description = "셀안의 특정 단어만 찾아 서식을 지정할 때 사용"
Attribute sb특정단어서식변경.VB_ProcData.VB_Invoke_Func = " \n14"
   '--// frm서식 폼에 RefEdit 컨트롤을 사용하기때문에
   '      해당 폼은 vbModeless로 사용할 수 없음
   frm서식.Show
End Sub


