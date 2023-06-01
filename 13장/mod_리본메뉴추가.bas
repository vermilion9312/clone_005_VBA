Attribute VB_Name = "mod_리본메뉴추가"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  리본 메뉴를 추가한 후 해당 명령을 클릭할 때 실행할
'            프로시저를 작성
'------------------------------------------------------------------------------------------
Option Explicit

Sub RibbonControl_Click(button As Office.IRibbonControl)
   Select Case button.ID
      Case "Button1": Call sbMsg_Time
      Case "Button2": Call sbMsg_Date
      
      Case Else: Call btnMsg(button.ID)
      
   End Select
End Sub

Sub sbMsg_Time()
   MsgBox "현재 시간 :" & Time, vbInformation, "리본 메뉴 연습"
End Sub

Sub sbMsg_Date()
   MsgBox "현재 날짜 :" & Date, vbInformation, "리본 메뉴 연습"
End Sub

Sub btnMsg(btnId As String)
   MsgBox btnId & "에 대한 처리를 선언하지 않았습니다.", vbCritical
End Sub
