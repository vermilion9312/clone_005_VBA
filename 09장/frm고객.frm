VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm고객 
   Caption         =   "고객정보 조회"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
   OleObjectBlob   =   "frm고객.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm고객"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbo정렬항목_Click()
   Dim keyCol As Integer
   Dim rngT As Range
   Set rngT = Range("nm고객_출력").CurrentRegion
   keyCol = Application.Match(Me.cbo정렬항목, rngT.Rows(1), 0) + rngT.Column - 1
   '--// 정렬 메서드는 2007 이후 방식임 : 2003 이하에선 오류 발생
   With rngT.Parent.Sort
      .SortFields.Clear
      .SortFields.Add Key:=rngT.Parent.Cells(rngT.Row, keyCol), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .SetRange rngT
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .Apply
   End With
End Sub

Private Sub cmd검색_Click()
   Dim rngT As Range
   
   '--// <고객목록> 시트의 고객 중 성명에  'txt검색이름' 내용을 포함하는 자료만
   '      고급필터를 이용하여  <작업시트>에 출력
   Range("nm고객_조건").Cells(2, 1) = "*" & Me.txt검색이름 & "*"
   Range("tbl고객정보[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm고객_조건"), _
        CopyToRange:=Range("nm고객_출력"), Unique:=False
        
   '--// <작업시트>에 출력된 내용을 lst고객명단 목록 상자에 표시
   Set rngT = Range("nm고객_출력").CurrentRegion
   
   With Me.lst고객명단
      .ColumnCount = 5
      .ColumnWidths = "2cm;1.5cm;3cm;0cm;0cm"
      .ColumnHeads = True
      If rngT.Rows.Count = 1 Then
         .RowSource = ""
      Else
         .RowSource = rngT.Offset(1, 0).Resize(rngT.Rows.Count - 1).Address(External:=True)
      End If
      If .ListCount > 0 Then .ListIndex = .ListCount - 1
   End With
   
   '--// 정렬 항목으로 자료 정렬
   Call cbo정렬항목_Click
End Sub

Private Sub cmd삭제_Click()
   Dim iOK As Integer, R As Long
   
   iOK = MsgBox("현재 조회 중인 자료를 삭제하시겠습니까?", vbYesNo + vbQuestion, "삭제확인")
   If iOK = vbYes Then
      R = Application.Match(Me.txt고객코드, Range("tbl고객정보[고객코드]"), 0)
      If R > 0 Then
         R = R + Range("tbl고객정보[[#Headers],[고객코드]]").Row
         Sheets("고객목록").Rows(R).Delete Shift:=xlUp
         MsgBox "삭제가 완료되었습니다.", vbInformation
      End If
      
      '--// 삭제후 목록 상자 내용 다시 표시
      Call cmd검색_Click
   End If
End Sub

Private Sub cmd수정_Click()
   Call sb컨트롤잠금(False)
End Sub

Private Sub cmd신규_Click()
   Call sb컨트롤잠금(False)
   Call sb컨트롤내용비우기
   
   '--// 코드를 새로 부여. 코드를 자동 부여하기 위해 <작업시트> A2 셀에 미리
   '     배열수식을 이용하여 고객코드 중 최대 숫자를 기억하고 있음
   Me.txt고객코드 = "S" & Format(Range("nmMax코드") + 1, "00000")
   Me.txt성명.SetFocus
   Call sb버튼표시(False)
End Sub

Private Sub cmd저장_Click()
   Dim R As Long
   
   If Me.txt고객코드 = "" Or Me.txt성명 = "" Then
      MsgBox "고객코드와 성명을 입력하세요.", vbCritical
      Exit Sub
   End If
   '--// <고객목록> 시트에서 txt고객코드 컨트롤의 고객코드가 몇번째 행에 위치하는지
   '    확인. 못찾은 경우 신규 등록하기 위해 현재 자료의 가장 마지막 행 다음 행을 반환
   If Application.CountIf(Range("tbl고객정보[고객코드]"), Me.txt고객코드) = 0 Then
      R = Range("tbl고객정보[고객코드]").End(xlDown).Row + 1
   Else
      R = Application.Match(Me.txt고객코드, Range("tbl고객정보[고객코드]"), 0)
      R = R + Range("tbl고객정보[[#Headers],[고객코드]]").Row
   End If
   
   With Sheets("고객목록")
      .Cells(R, 1) = Me.txt고객코드
      .Cells(R, 2) = Me.txt성명
      .Cells(R, 3) = Me.txt소속
      .Cells(R, 4) = Me.txt연락처
      .Cells(R, 5) = Me.txt주소
   
   MsgBox "저장되었습니다.", vbInformation
   
   '--// 등록된 내용이 목록 상자에 반영되도록 <검색> 버튼을 클릭한 것처럼 동작
   Call cmd검색_Click
   '--// 최근 수정/등록한 사람이 표시되도록 목록 값을 지정
   Me.lst고객명단.Text = .Cells(R, 1)
   End With
End Sub

Private Sub cmd취소_Click()
   Call lst고객명단_Click
End Sub

Private Sub lst고객명단_Click()
   With Me.lst고객명단
      If .ListIndex >= 0 Then
         Me.txt고객코드 = .Column(0)
         Me.txt성명 = .Column(1)
         Me.txt소속 = .Column(2)
         Me.txt연락처 = .Column(3)
         Me.txt주소 = .Column(4)
      Else
         Call sb컨트롤내용비우기
      End If
   End With
   
   Call sb컨트롤잠금(True)
End Sub


Private Sub txt검색이름_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = 13 Then
      Call cmd검색_Click
      With Me.txt검색이름
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
      End With
      KeyCode = 0       '--// 엔터키를 무효 시켜서 다른 컨트롤로 커서가 이동되지 않도록 함
   End If
End Sub


Private Sub UserForm_Initialize()
   With Me.cbo정렬항목
      .Clear
      .AddItem "고객코드"
      .AddItem "성명"
      .AddItem "소속"
      .AddItem "연락처"
      .AddItem "주소"
      .Text = "고객코드"
   End With
   
   Call cmd검색_Click
End Sub

'-------------------------------------------------------------
' 수정과 신규 등록 상태에서만 텍스트 상자가 사용할 수 있도록 처리
' 컨트롤 특수효과(SpecialEffect)도 상황에 따라 달라지게 함
'-------------------------------------------------------------
Sub sb컨트롤잠금(bLock As Boolean)
   Dim ctrNM
   Dim i As Integer
   ctrNM = Array("txt성명", "txt소속", "txt연락처", "txt주소")
   
   For i = LBound(ctrNM) To UBound(ctrNM)
      Me.Controls(ctrNM(i)).Locked = bLock
      Me.Controls(ctrNM(i)).SpecialEffect = IIf(bLock, 3, 2)
   Next
   
   Call sb버튼표시(bLock)
End Sub

Sub sb컨트롤내용비우기()
   Me.txt고객코드 = ""
   Me.txt성명 = ""
   Me.txt소속 = ""
   Me.txt연락처 = ""
   Me.txt주소 = ""
End Sub

'-------------------------------------------------------------
' 수정과 신규 등록 상태에서만 저장/취소 버튼이 표시되고
' 그 이외에는 숨기기하기 위한 처리
'-------------------------------------------------------------
Sub sb버튼표시(bShow As Boolean)
   Me.cmd삭제.Visible = bShow
   Me.cmd수정.Visible = bShow
   Me.cmd신규.Visible = bShow
   Me.cmd저장.Visible = Not bShow
   Me.cmd취소.Visible = Not bShow
End Sub

