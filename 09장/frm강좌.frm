VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm강좌 
   Caption         =   "강좌조회 및 예약"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "frm강좌.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm강좌"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
   Dim i As Long
   With Me.lst예약명단
      For i = 0 To .ListCount - 1
         .Selected(i) = chkAll.Value
      Next
   End With
End Sub

Private Sub cmd검색_Click()
   Dim rngT As Range
   
   With Range("nm강좌_조건")
      .Cells(2, 1) = "*" & Me.txt검색이름 & "*"
      .Cells(2, 2) = IIf(Me.txtFrom = "", "", ">=") & Me.txtFrom
      .Cells(2, 3) = IIf(Me.txtTo = "", "", "<=") & Me.txtTo
   End With
   
   Range("tbl개설강좌[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm강좌_조건"), _
        CopyToRange:=Range("nm강좌_출력"), Unique:=False
        
   Set rngT = Range("nm강좌_출력").CurrentRegion
   
   With Me.lst강좌명단
      .ColumnCount = 8
      .ColumnWidths = "1.5 cm;2.5 cm;4 cm;3 cm;2.5 cm;2.5 cm;1 cm;1 cm"
      .ColumnHeads = True
      If rngT.Rows.Count = 1 Then
         .RowSource = rngT.Offset(1, 0).Address(External:=True)
      Else
         .RowSource = rngT.Offset(1, 0).Resize(rngT.Rows.Count - 1).Address(External:=True)
      End If
      If .ListCount > 0 Then .ListIndex = .ListCount - 1
   End With
End Sub


Private Sub cmd명단출력_Click()
   Dim R As Long, Cnt As Long, i As Long
   Dim rngT As Range
   
   With Sheets("출력-강좌별")
      .Range("C3") = Me.txt강좌명 & " (" & Me.txt강좌코드 & ")"
      .Range("C4") = Format(Me.txt일자, "yy-mm-dd(aaa)")
      .Range("C5") = Me.txt장소
      .Range("G3") = Me.lst강좌명단.Column(5)
      Set rngT = .Range("A8")
      '--// 8행 이하 사용한 영역 삭제
      .Range(rngT.Offset(1, 0), rngT.SpecialCells(xlLastCell)).Clear
   End With
      
   With Me.lst예약명단
      rngT.EntireRow.Copy
      With rngT.Offset(1, 0).Resize(.ListCount - 1).EntireRow
            .PasteSpecial Paste:=xlPasteFormats
            .PasteSpecial Paste:=xlPasteFormulas
      End With
      Application.CutCopyMode = False
      
      For i = 0 To .ListCount - 1
         rngT.Offset(i, 0) = i + 1
         rngT.Offset(i, 1) = .List(i, 0)  '--// 고객코드
         rngT.Offset(i, 3) = .List(i, 1)  '--// 고객명
         rngT.Offset(i, 6) = Format(.List(i, 2), "yy-mm-dd(aaa)") '--//예약일
         rngT.Offset(i, 7) = .List(i, 3)  '--// 비고
         rngT.Offset(i, 8) = .List(i, 4)  '--// 참석여부
      Next
      '--// 인쇄 영역 다시 지정
      Sheets("출력-강좌별").PageSetup.PrintArea = "$A$1:$I$" & (rngT.Row + .ListCount)
      Sheets("출력-강좌별").PrintOut preview:=True
   End With
End Sub

Private Sub cmd삭제_Click()
   Dim i As Long, R As Long, iOK As Integer
   Dim strKey
  
   iOK = MsgBox("선택하신 자료들을  삭제하시겠습니까?", vbYesNo + vbQuestion, "삭제확인")
   If iOK = vbYes Then
      With Me.lst예약명단
         For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
               strKey = Me.txt강좌코드 & .List(i, 0)
               If Application.CountIf(Range("tbl예약현황[Key]"), strKey) > 0 Then
                  R = Application.Match(strKey, Range("tbl예약현황[Key]").EntireColumn, 0)
                  Sheets("예약현황").Rows(R).Delete Shift:=xlUp
               End If
            End If
         Next
      End With
      
      MsgBox "삭제가 완료되었습니다.", vbInformation
      Call sb예약명단출력
   End If
End Sub

Private Sub cmd신규_Click()
   Dim i As Long, R As Long
   Dim strKey
   Application.DisplayAlerts = False
   
   With Me.lst고객명단
      For i = 0 To .ListCount - 1
         If .Selected(i) Then
            strKey = Me.txt강좌코드 & .List(i, 0)
            If Application.CountIf(Range("tbl예약현황[Key]"), strKey) = 0 Then
               If Range("tbl예약현황[Key]").Rows.Count = 1 Then
                  R = Range("tbl예약현황[Key]").Row + 1
               Else
                  R = Range("tbl예약현황[Key]").End(xlDown).Row + 1
               End If
               
               With Sheets("예약현황")
                  .Cells(R, 2) = Me.txt강좌코드
                  .Cells(R, 3) = Me.lst고객명단.List(i, 0)
                  .Cells(R, 5) = Date
               End With
            End If
         End If
      Next
   End With
   
   Call sb예약명단출력
   Application.DisplayAlerts = True

End Sub


Private Sub lst강좌명단_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Call sb예약명단출력
   Call sb고객명단출력
   Me.MultiPage1.Value = 1
End Sub

Sub sb예약명단출력()
   Dim rngT As Range
   
   '--//  [강좌조회] 페이지에서 선택한 강좌정보 표시
   With Me.lst강좌명단
      Me.txt강좌코드 = .Column(0)
      Me.txt강좌명 = .Column(2)
      Me.txt일자 = .Column(1)
      Me.txt장소 = .Column(3)
   End With
   
   '--// [예약현황] 시트의 내용 중 선택한 강좌 코드의 정보만 고급필터로 추출
   Range("nm예약_조건").Cells(2, 1) = Me.txt강좌코드
   Range("tbl예약현황[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm예약_조건"), _
        CopyToRange:=Range("nm예약_출력"), Unique:=False
        
   '--// 고급필터 출력 내용을 예약명단 목록 상자에 표시
   Set rngT = Range("nm예약_출력").CurrentRegion
   
   With Me.lst예약명단
      .ColumnCount = 5
      .ColumnWidths = "2 cm;1.5 cm;2.5 cm;2 cm;1 cm"
      .ColumnHeads = True
      .MultiSelect = fmMultiSelectExtended
      If rngT.Rows.Count = 1 Then
         .RowSource = ""
      Else
         .RowSource = rngT.Offset(1, 0).Resize(rngT.Rows.Count - 1).Address(External:=True)
      End If
      If .ListCount > 0 Then .ListIndex = .ListCount - 1
   End With
End Sub

Sub sb고객명단출력()
   Dim rngT As Range
   '--// [고객목록] 시트의 고객 명단을 '고객명단' 목록 상자에 출력
   Set rngT = Range("tbl고객정보[#All]")
   
   With Me.lst고객명단
      .ColumnCount = 5
      .ColumnWidths = "2cm;1.5cm;3cm;0cm;0cm"
      .ColumnHeads = True
      .MultiSelect = fmMultiSelectExtended
      If rngT.Rows.Count = 1 Then
         .RowSource = ""
      Else
         .RowSource = rngT.Offset(1, 0).Resize(rngT.Rows.Count - 1).Address(External:=True)
      End If
      If .ListCount > 0 Then .ListIndex = .ListCount - 1
   End With
End Sub

Private Sub MultiPage1_Change()
   If Me.MultiPage1.Value = 1 And Me.txt강좌코드 = "" Then
      Call sb예약명단출력
      Call sb고객명단출력
   End If
End Sub

Private Sub UserForm_Initialize()
   Me.MultiPage1.Value = 0
   Call cmd검색_Click
   chkAll = False
End Sub
