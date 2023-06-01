Attribute VB_Name = "mod특정시간실행"
Option Explicit
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  1초마다 한번씩 실행되어 B2 셀 값을 1씩 증가시키는 매크로
'   주의 사항: 파일을 닫을때 자동기능을 종료해야함
'------------------------------------------------------------------------------------------
Dim setTime    '--// 실행할 시간 기억

Sub sbTimer_Start()
Attribute sbTimer_Start.VB_Description = "1초마다 한번씩 실행되어 B2 셀 값을 1씩 증가시키는 매크로"
Attribute sbTimer_Start.VB_ProcData.VB_Invoke_Func = " \n17"
    setTime = Now + TimeValue("00:00:01")         '--//이벤트 실행할 시간을 현재 시간에 1초를 더해 기억
    Application.OnTime setTime, "sbTimer_Start"    '--//setTime 시간에 매크로 실행
    
    Range("B2") = Range("B2").Value + 1      '--// 값을 1 초씩 증가
    Range("B2").NumberFormatLocal = "#,##0초"
End Sub


Sub sbTimer_Stop()
Attribute sbTimer_Stop.VB_Description = "sbTimer_Start 자동 실행 기능을 해제"
Attribute sbTimer_Stop.VB_ProcData.VB_Invoke_Func = " \n17"
    On Error Resume Next
    '--// setTime 시간에 실행할 sbTimer_Start 프로시저 해제
    Application.OnTime setTime, "sbTimer_Start", , False
End Sub

Sub sbTimer_Clear()
    Range("B2").Value = 0
    Call sbTimer_Start
End Sub






