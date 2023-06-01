VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm서식 
   Caption         =   "특정 단어 서식만 변경하기 // 길벗 출판사"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   OleObjectBlob   =   "frm서식.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm서식"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  셀안의 특정 단어만 찾아 서식을 지정할 때 사용
'   함께 사용해야할 자료 : modPictureSaveAPI , frm서식
'------------------------------------------------------------------------------------------

Option Explicit
Option Compare Text  '영문자 비교시 대소문자를 구분하지 않고 비교

Private Sub cmd글꼴서식_Click()
   Dim wkA As Worksheet
   
   Application.ScreenUpdating = False
   
   If Me.txt찾을단어.Text = "" Then
      ThisWorkbook.Sheets("#Color").Range("A1").Value = "변경할 글꼴 서식"
   Else
      ThisWorkbook.Sheets("#Color").Range("A1").Value = Me.txt찾을단어
   End If
   
   Set wkA = ActiveSheet   '--// 현재 선택하고 있는 시트를 기억
   '--// 글꼴 서식을 지정하기 위해, '#Color' 시트의 A1 셀을 선택함
   ThisWorkbook.Sheets("#Color").Visible = True
   ThisWorkbook.Sheets("#Color").Select
   Range("A1").Select
   '--// Dialogs(xlDialogFontProperties)를 이용하면 '글꼴 서식' 대화상자가 나타나고
   '     현재 선택한 영역에 자동으로 선택한 글꼴 서식이 지정됨
   Application.Dialogs(xlDialogFontProperties).Show
   wkA.Select     '--// 처음 선택했던 시트가 다시 선택
   
   Call 그림으로저장
   
   Application.ScreenUpdating = True
End Sub

Sub 그림으로저장()
   Dim sFile As String
   '--// 글꼴 서식이 적용된 A1 셀 영역을 그림으로 저장
   sFile = ThisWorkbook.Path & "\temp.jpg"
   dhSavePic ThisWorkbook.Sheets("#Color").Range("A1"), sFile
    
   '--// 이미지 컨트롤에 저장한 그림을 표시
   Me.Image1.Picture = LoadPicture(sFile)
   Me.Image1.PictureSizeMode = fmPictureSizeModeZoom
   Kill sFile
   ThisWorkbook.Sheets("#Color").Visible = False
End Sub

Private Sub cmd닫기_Click()
   Unload Me
End Sub

Private Sub UserForm_Initialize()
   '--// RGB 값을 가져오기 위한 시트 삽입
   Dim wkA As Worksheet
   Dim k As Integer
   
   Me.RefEdit1 = Selection.Address
   
   On Error Resume Next
   Set wkA = ThisWorkbook.Sheets("#Color")
   On Error GoTo 0
   
   If wkA Is Nothing Then
      Set wkA = ThisWorkbook.Sheets.Add
      wkA.Name = "#Color"
   End If
   With wkA.Range("A1")
      .Value = "변경할 글꼴 서식"
      .ColumnWidth = 30
      .RowHeight = 60
   End With

   Call 그림으로저장
End Sub

Private Sub cmd확인_Click()
   Dim rngFind As Range, rngFirst As Range, rngWork As Range, rngK As Range
   Dim iST As Long, iLen As Long, i As Long
   iLen = Len(Me.txt찾을단어)
   
   If Me.RefEdit1 = "" Then
      MsgBox "작업 범위를 선택하세요.", vbCritical, "범위선택 오류"
      Exit Sub
   End If
   
  
   Set rngWork = Range(Me.RefEdit1)    '--// 단어를 찾을 범위를 rngWork 변수에 저장
   Set rngK = ThisWorkbook.Sheets("#Color").Range("A1")  '--// 적용할 서식이 적용된 셀을 rngK 변수에 저장
   If rngWork.Cells.Count = 1 Then
      '--// 선택 영역이 한 셀인 경우 해당 셀에 찾는 내용이 있는지 확인
      If InStr(rngWork, Me.txt찾을단어) > 0 Then
         Set rngFirst = rngWork
      End If
   Else
      '--// 선택 영역이 여러 셀인 경우 '찾기 기능'(Find)을 이용해 찾음
      Set rngFirst = rngWork.Find(What:=Me.txt찾을단어, After:=rngWork.Range("A1"), _
         LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
         SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
   End If
   '--// 첫번째 찾은 셀을 rngFirst 변수에 기억시켜 두어, Find 검색의 결과가
   '   rngFirst 변수와  같은 셀이될 때 찾기 기능을 종류함
   Set rngFind = rngFirst
   
   If rngFind Is Nothing Then
      MsgBox "해당 자료를 찾을 수 없습니다.", vbCritical, "자료 없음"
      Exit Sub
   End If
   
   '--// 지정된 검색 영역(rngWork)에서 '다음 찾기'를 이용해 해당 단어를 검색하여
   '     서식을 지정
   i = 0
   Do
      i = i + 1
      iST = InStr(rngFind, Me.txt찾을단어)
      With rngFind.Characters(Start:=iST, Length:=iLen).Font
          .Name = rngK.Font.Name
          .FontStyle = rngK.Font.FontStyle
          .Size = rngK.Font.Size
          .Italic = rngK.Font.Italic
          .Bold = rngK.Font.Bold
          .Color = rngK.Font.Color
          .Underline = rngK.Font.Underline
          .Strikethrough = rngK.Font.Strikethrough
          .Superscript = rngK.Font.Superscript
          .Subscript = rngK.Font.Subscript
          .OutlineFont = rngK.Font.OutlineFont
          .Shadow = rngK.Font.Shadow
          .TintAndShade = rngK.Font.TintAndShade
          .ThemeFont = rngK.Font.ThemeFont
          .TintAndShade = rngK.Font.TintAndShade
      End With
    
      Set rngFind = rngWork.FindNext(After:=rngFind)
      If rngFind Is Nothing Then Exit Do

    Loop While rngFind.Address <> rngFirst.Address
    
    MsgBox i & "번을 찾아 서식을 변경했습니다.", vbInformation, "작업 완료"
    
End Sub

