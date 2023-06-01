Attribute VB_Name = "modMain"
'-----------------------------------------------------------------
'  기   능 : 사용자 정의 폼을 실행하기 위한 매크로
'  작성자 :  길벗 엑셀 매크로와 VBA
'-----------------------------------------------------------------
Option Explicit

Sub 그룹별시트분리()
Attribute 그룹별시트분리.VB_ProcData.VB_Invoke_Func = " \n14"
   UserForm1.Show
End Sub

'-------------------------------------------------
'  리본 메뉴에 명령을 추가하여 사용하는 경우
'  customUI.xml과 .rels 파일의 수정이 필요함
'  sbChooseMacro는 리본 메뉴의 버튼과 연결되어 실행됨
'-------------------------------------------------
Sub sbChooseMacro(button As IRibbonControl)
  Select Case button.ID
    Case "customButton1"
      UserForm1.Show
    Case "customButton2"
      Call 배송업체별그룹분류
    Case "customButton3"
      Call 열제목지정그룹분류
    Case Else
      MsgBox "해당 기능과 연결된 작업이 없습니다.", vbInformation, "안내"
  End Select
End Sub

