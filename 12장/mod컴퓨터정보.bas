Attribute VB_Name = "mod컴퓨터정보"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : 사용중인 컴퓨터의 컴퓨터명과 로그인 사용자 ID를 표시
'------------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------------------
'   로그인한 사용자 ID와 컴퓨터 이름, 유령 문자 제거 API 함수
'   엑셀 2010 이상에서 엑셀 버전이 32bit일 때와 64bit일 때 다르게 처리
'------------------------------------------------------------------------------------------
'  VBA7은 2010 이상의 VBA를 의미로 이전 버전의 VBA 코드인지를 비교할 때 사용
'  Win64는 32bit인지 64bit 인지 구분할 때 사용
#If VBA7 And Win64 Then      '--// 64bit
   Private Declare PtrSafe Function GetComputerName Lib "kernel32" _
            Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
   Private Declare PtrSafe Function GetUserName Lib "advapi32" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
   Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

#Else                '--//32Bit
   Private Declare Function GetComputerName Lib "kernel32" _
            Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function GetUserName Lib "advapi32" _
            Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
   Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
#End If

'------------------------------------------------------------------------------------------
'    컴퓨터이름을 반환
'------------------------------------------------------------------------------------------
Function fn컴퓨터명()
Attribute fn컴퓨터명.VB_Description = "컴퓨터이름을 반환"
Attribute fn컴퓨터명.VB_ProcData.VB_Invoke_Func = " \n14"
      Dim tmp As String
      Application.Volatile
     
      tmp = Space$(256)
      
   GetComputerName tmp, 256
   fn컴퓨터명 = Left$(tmp, lstrlenW(StrPtr(tmp)))
End Function

'------------------------------------------------------------------------------------------
' 컴퓨터 로그인 사용자 ID 반환
'------------------------------------------------------------------------------------------
Function fn컴퓨터사용자()
Attribute fn컴퓨터사용자.VB_Description = "컴퓨터 로그인 사용자 ID 반환"
Attribute fn컴퓨터사용자.VB_ProcData.VB_Invoke_Func = " \n14"
      Dim tmp As String
      Application.Volatile
      
      tmp = Space$(256)
      
   GetUserName tmp, 256
   fn컴퓨터사용자 = Left$(tmp, lstrlenW(StrPtr(tmp)))
End Function



