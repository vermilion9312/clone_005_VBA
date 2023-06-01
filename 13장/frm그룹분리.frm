VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm그룹분리 
   Caption         =   "그룹별 시트 분리 //길벗 출판사"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   OleObjectBlob   =   "frm그룹분리.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm그룹분리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  표 형태의 데이트를 지정하면, 첫 행을 필드명으로 인식한 후
'            해당 필드에 속한 자료들을 분리하여 별도의 시트로 복사
'------------------------------------------------------------------------------------------

Option Explicit
Dim rngTable As Range

'---------------------------------------------------------------
' 폼 실행시 현재 셀에 내용이 있는 경우 현재 셀 영역을
' 범위로 자동 지정
'---------------------------------------------------------------
Private Sub UserForm_Initialize()
   On Error Resume Next
   lblMsg.Caption = "작업 범위와 분리 열을 선택한 후 <실행>을 클릭하세요."
   If ActiveCell <> "" Then
      Me.RefEdit1.Text = "'" & ActiveSheet.Name & "'!" & ActiveCell.CurrentRegion.Address
   End If
   On Error GoTo 0
End Sub

Private Sub RefEdit1_Change()
   lblMsg.Caption = "작업 범위와 분리 열을 선택한 후 <실행>을 클릭하세요."
End Sub

Private Sub cboCol_Change()
   lblMsg.Caption = "작업 범위와 분리 열을 선택한 후 <실행>을 클릭하세요."
End Sub

Private Sub cmd닫기_Click()
   Unload Me
End Sub

'---------------------------------------------------------------
' 작업 범위가 지정된 후 컨트롤을 빠져나갈 때
' cboCol 컨트롤의 목록을 작업 범위 첫 행의 내용으로  다시 표시
'---------------------------------------------------------------
Private Sub RefEdit1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
   Dim rngK As Range
   If Trim(Me.RefEdit1.Text) = "" Then Exit Sub
   Set rngTable = Range(Me.RefEdit1)
   If rngTable.Areas.Count > 1 Then
      MsgBox "셀 영역은 연속된 단일 영역이여야 합니다.", vbCritical, "다중 영역 오류"
      Exit Sub
   End If
   
   Me.cboCol.Clear
   For Each rngK In rngTable.Rows(1).Cells
      Me.cboCol.AddItem rngK
   Next
End Sub

'---------------------------------------------------------------
' 작업 실행
'---------------------------------------------------------------
Private Sub cmd실행_Click()
   Dim GroupList  As New Collection
   Dim varK As Variant
   Dim R As Long, ColNo As Long
   Dim rngK As Range
   Dim wkB As Workbook, sh As Worksheet
   
   If Me.RefEdit1.Text = "" Or Me.cboCol.Text = "" Then Exit Sub
   
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   ColNo = Me.cboCol.ListIndex + 1    '--// 열의 위치번호
   
   '--// GroupList 변수에 지정할 열(ColNo)의 자료들을 중복없이 하나씩 저장(배열 변수 형태가 됨)
   On Error Resume Next
   '--// 지정한 열에서 첫 셀은 필드명이기때문에 제외하고 두번째 셀부터 시작
   For R = 2 To rngTable.Columns(ColNo).Cells.Count
      Set rngK = rngTable.Columns(ColNo).Cells(R, 1)
      If TypeName(rngK.Value) = "Date" Then
         '--// 셀내용이 날짜인 경우에는 셀 서식 상관없이 값으로 중복 체크
         GroupList.Add Item:=rngK.Value, Key:="D" & rngK.Value
      Else
         '--// 자동 필터에서 날짜를 제외한 값들은 모두 서식이 적용된 텍스트로 인식
         GroupList.Add Item:=rngK.Text, Key:=rngK.Text
      End If
   Next
   On Error GoTo 0
   
On Error GoTo End_Rtn
   If GroupList.Count = 0 Then Exit Sub
   Me.MousePointer = fmMousePointerHourGlass '--// 폼에서 마우스 포인터 모양을 모래시계로 변경
   
   '--// 새 워크북 추가
   Set wkB = Workbooks.Add
   
   '--// 저장된 GroupList의 내용으로 자동필터한후 필터 결과를 복사하여 새 워크시트에 붙여넣기하여 추출
   For Each varK In GroupList
      '--// 진행 메시지를 표시
      lblMsg.Caption = "【" & varK & "】에 대한 분리 작업을 진행 중입니다. 잠시만 기다리세요."
      Me.Repaint
      
      '--// 필드의 데이터 종류가 날짜형일 때는 년/월로 구분되게 지정
      If TypeName(rngTable.Cells(2, ColNo).Value) = "Date" Then
         rngTable.AutoFilter Field:=ColNo, Operator:=xlFilterValues, Criteria2:=Array(2, varK)
      Else
         rngTable.AutoFilter Field:=ColNo, Criteria1:=varK
      End If
      rngTable.Copy
      
      wkB.Sheets.Add After:=Sheets(Sheets.Count)
      ActiveSheet.Name = varK
      ActiveSheet.Paste
      Selection.Columns.AutoFit
   Next
   '--// 불필요한 시트 삭제
   For Each sh In wkB.Sheets
      If sh.UsedRange.Address = "$A$1" Then sh.Delete
   Next
    
   '--// 자동필터가 되어있는 상태에서 필터 조건 모두 지우기
   If rngTable.Parent.AutoFilter.FilterMode Then rngTable.Parent.AutoFilter.ShowAllData
   wkB.Activate
   
End_Rtn:
   Application.CutCopyMode = False
   Me.MousePointer = fmMousePointerDefault  '--// 폼에서 마우스 포인터 모양을 기본으로 변경
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   If Err.Number = 0 Then
      lblMsg.Caption = "작업이 완료되었습니다."
      MsgBox "작업이 완료되었습니다.", vbInformation, "완료"
   Else
      lblMsg.Caption = "작업 중 다음과 같은 오류가 발생했습니다."
      MsgBox "작업 중 다음과 같은 오류가 발생했습니다." & vbCrLf & _
               Err.Description, vbCritical, "오류"
   End If
End Sub
