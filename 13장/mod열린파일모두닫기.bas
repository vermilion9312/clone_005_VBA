Attribute VB_Name = "mod열린파일모두닫기"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  현재 문서를 제외한 문서가 열려있는 경우 해당 문서의 저장 여부를 확인한 후 닫는 처리
'------------------------------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------------------------------
'  현재 문서를 제외한 문서가 열려있는 경우 해당 문서의 저장 여부를 확인한 후 닫는 처리
'  처리과정 메시지를 출력할 폼 및 무조건 확인 없이 저장하고 종료할 지 여부 매개변수 처리
'--------------------------------------------------------------------------------------
Sub sbClose()
Attribute sbClose.VB_Description = "현재 문서를 제외한 문서가 열려있는 경우 해당 문서의 저장 여부를 확인한 후 닫는 처리"
Attribute sbClose.VB_ProcData.VB_Invoke_Func = " \n17"
    Dim i As Integer
    Dim K As Workbook
    Dim strMsg As String, bSave As Boolean
    
   If Workbooks.Count = 1 Then Exit Sub
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
On Error GoTo End_Rtn
   i = MsgBox("현재 문서를 제외한 열려진 파일이 【총" & Workbooks.Count - 1 & "개】입니다." & vbCrLf & _
       "이 문서들을 닫고 작업해야 합니다. 저장하고 닫을까요?" & vbCrLf & _
       "【예】-저장하고 닫기" & vbCr & "【아니오】-저장하지않고 닫기" & vbCr & "【취소】-작업취소", vbQuestion + vbYesNoCancel, "파일 닫기 확인")
   
   Select Case i
       Case vbYes
           bSave = True: strMsg = "현재 열려진 파일을 저장하고 닫는 중입니다. "
       Case vbNo
           bSave = False: strMsg = "현재 열려진 파일을 저장하고 닫는 중입니다. "
       Case Else
         End
   End Select
    
   For Each K In Workbooks
        If K.Name <> ThisWorkbook.Name Then
            If K.ReadOnly Then   '--// 읽기 전용으로 열린 파일을 체크
                K.Close False
            Else
                K.Close bSave
            End If
        End If
   Next
    
End_Rtn:
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   If Err.Number = 0 Then
      MsgBox "작업을 정상적으로 완료했습니다.", vbInformation, "작업완료"
   Else
      MsgBox "작업 중 다음과 같은 오류가 발생했습니다." & vbCrLf & _
               Err.Description, vbCritical, "오류"
   End If
End Sub


