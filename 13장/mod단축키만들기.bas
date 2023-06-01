Attribute VB_Name = "mod단축키만들기"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 : Application.OnKey 메서드를 이용하여 특정 키 조합에 기능을 설정하기
'   전제조건 : 해당 기능이 자동으로 실행되도록 할 때는
'                 ThisWorkbook(현재_통합_문서) 에서 'Workbook_Open' 이벤트를 이용해
'                 실행되도록 해야 함
'------------------------------------------------------------------------------------------
Option Explicit

Sub sb키기능설정()
Attribute sb키기능설정.VB_Description = "Application.OnKey 메서드를 이용하여 특정 키 조합에 기능을 설정하기"
Attribute sb키기능설정.VB_ProcData.VB_Invoke_Func = " \n17"
    Application.OnKey "+{F9}", "sb메시지"
    MsgBox "<Shift>+<F9> 키에 대한 기능키 설정이 설정되었습니다."
End Sub

Sub sb키기능삭제()
Attribute sb키기능삭제.VB_Description = "Application.OnKey 메서드를 이용하여 특정 키 조합에 기능을 해제하기"
Attribute sb키기능삭제.VB_ProcData.VB_Invoke_Func = " \n17"
    Application.OnKey "+{F9}"
    MsgBox "<Shift>+<F9> 키에 대한 기능키 설정이 해제되었습니다."
End Sub

Sub sb메시지()
   MsgBox Time, vbInformation, "현재 시간"
End Sub

