Attribute VB_Name = "mod시트관리"
Option Explicit
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  현재 열어놓은 파일들의 목록과 파일별 시트 목록을 표시
'           시트의 위치, 정렬 방법, 추가, 삭제, 이름 변경, 숨기기 등이 가능
'   함께 사용할 모듈 : frm시트관리.frm
'------------------------------------------------------------------------------------------

Sub sb시트관리()
Attribute sb시트관리.VB_Description = "현재 열어놓은 파일들의 목록과 파일별 시트 목록을 표시한 후 시트의 위치, 정렬 방법, 추가, 삭제, 이름 변경, 숨기기 등이 가능"
Attribute sb시트관리.VB_ProcData.VB_Invoke_Func = " \n17"
   '--// 시트 순서를 위로/아래로 이동하는(Move) 작업시
   '  폼이 vbModeless로 표시되면 오류가 발생하므로
   '  선택하지 않고 사용해야 함
    frm시트관리.Show
End Sub


