Attribute VB_Name = "mod문자선택적추출"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  문자열에서 구분을 1~4까지로 지정하여
'            숫자(구분:1), 영문자(구분:2), 한글(구분:3), 기타문자(구분:4)
'            만 추출하여 반환
'------------------------------------------------------------------------------------------
Function fn문자추출(문자열 As String, 구분 As Integer) As String
Attribute fn문자추출.VB_Description = "문자열에서 구분을 1~4까지로 지정하여 \n숫자(구분:1), 영문자(구분:2), 한글(구분:3), 기타문자(구분:4)만 추출하여 반환"
Attribute fn문자추출.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim i As Integer
  Dim K As String
  '--// 숫자, 영문, 한글, 기타 문자를 저장할 변수
  Dim NumStr As String, EngStr As String, HanStr As String, EtcStr As String  '기타 글자들을 기억함
                                 
  For i = 1 To Len(문자열)
      K = Mid(문자열, i, 1)
      Select Case K
         Case "0" To "9"
           NumStr = NumStr & K
         Case "."
           NumStr = NumStr & K
         Case "A" To "Z"
           EngStr = EngStr & K
         Case "a" To "z"
           EngStr = EngStr & K
         Case "ㄱ" To "홓"    '한글은 'ㄱ'이 가장 작고 '홓'이 가장 큰 글자
           HanStr = HanStr & K
         Case Else
           EtcStr = EtcStr & K
      End Select
  Next
  
  Select Case 구분
      Case 1:          fn문자추출 = NumStr
      Case 2:          fn문자추출 = EngStr
      Case 3:          fn문자추출 = HanStr
      Case 4:          fn문자추출 = EtcStr
      Case Else:       fn문자추출 = "오류"
  End Select
End Function

