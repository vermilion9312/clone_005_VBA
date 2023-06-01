Attribute VB_Name = "mod인터넷연결"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 /  이동숙(bofb@naver.com) 】
'   기능 :  인터넷 연결 상태 확인
'            URL 주소를 입력하여 해당 사이트 접속 가능 여부 판단
'            URL 주소 생략시 인터넷 연결 가능/불가만 체크
'------------------------------------------------------------------------------------------

Option Explicit

Private Const FLAG_ICC_FORCE_CONNECTION = &H1
#If VBA7 And Win64 Then      '--// 64bit
      Private Declare PtrSafe Function InternetCheckConnection Lib "wininet.dll" _
            Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, _
            ByVal dwflags As Long, ByVal dwReserved As Long) As Long
      Private Declare PtrSafe Function InternetAttemptConnect Lib "wininet" _
            (ByVal dwReserved As Long) As Long

#Else
      Private Declare Function InternetCheckConnection Lib "wininet.dll" _
            Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, _
            ByVal dwflags As Long, ByVal dwReserved As Long) As Long
      Private Declare Function InternetAttemptConnect Lib "wininet" _
            (ByVal dwReserved As Long) As Long

#End If

Function fn인터넷상태(Optional URL As String)
Attribute fn인터넷상태.VB_Description = "인터넷 연결 상태 확인\nURL 주소를 입력하여 해당 사이트 접속 가능 여부를 [정상연결]/[접속불가]로 표시\nURL 주소 생략시 인터넷 연결 가능 여부를 [연결가능]/[연결불가]로 표시"
Attribute fn인터넷상태.VB_ProcData.VB_Invoke_Func = " \n14"
   If InternetAttemptConnect(0) = 0 Then
      If URL = "" Then
         fn인터넷상태 = "연결가능"
      ElseIf InternetCheckConnection(URL, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
         fn인터넷상태 = "접속불가"
      Else
         fn인터넷상태 = "정상연결"
      End If
   Else
       fn인터넷상태 = "연결불가"
   End If
End Function



