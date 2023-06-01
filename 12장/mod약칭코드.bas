Attribute VB_Name = "mod약칭코드"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : 단어 첫문자로 약어 만들기
'------------------------------------------------------------------------------------------
Option Explicit

Function fn약칭코드(문자열 As Range)
Attribute fn약칭코드.VB_Description = "단어 첫문자로 약어 만들기"
Attribute fn약칭코드.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim str코드 As String, strTemp As String
   Dim i As Integer
   
   strTemp = Trim(문자열)
   
   Do
      i = i + 1
      str코드 = str코드 & Mid(strTemp, i, 1)
      i = InStr(i, strTemp, " ")
   Loop Until i = 0
   
   fn약칭코드 = UCase(str코드)
End Function

