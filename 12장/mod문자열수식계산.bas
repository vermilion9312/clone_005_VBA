Attribute VB_Name = "mod문자열수식계산"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  문자열 내용 중 숫자와 연산자 이외의 문자를 제외하고 계산
'------------------------------------------------------------------------------------------
Function fn문자열계산(문자열로된수식 As String)
Attribute fn문자열계산.VB_Description = "문자열 내용 중 숫자와 연산자 이외의 문자를 제외하고 계산"
Attribute fn문자열계산.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.Volatile
  
  Dim strResult As String, strTemp As String
  Dim i As Integer
  Const cOperator As String = "-+*/()^.%"
  
  For i = 1 To Len(문자열로된수식)
    strTemp = Mid(문자열로된수식, i, 1)
    If strTemp >= "0" And strTemp <= "9" Then
        strResult = strResult & strTemp
    ElseIf InStr(1, cOperator, strTemp) <> 0 Then
        strResult = strResult & strTemp
    End If
  Next
  
  fn문자열계산 = Application.Evaluate(strResult)
End Function

'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  문자열 내용을 무조건 계산
'            문자열 내용에 숫자, 연산자 이외에는 입력하면 0으로 계산
'------------------------------------------------------------------------------------------
Function fn단순문자열계산(문자열로된수식 As String)
Attribute fn단순문자열계산.VB_Description = "문자열 내용을 무조건 계산 문자열 내용에 숫자, 연산자 이외에는 입력하면 0으로 계산"
Attribute fn단순문자열계산.VB_ProcData.VB_Invoke_Func = " \n17"
  Application.Volatile
  
  단순문자열계산 = Application.Evaluate(문자열로된수식)
End Function


