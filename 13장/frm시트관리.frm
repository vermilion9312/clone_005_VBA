VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm시트관리 
   Caption         =   "시트관리 "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   OleObjectBlob   =   "frm시트관리.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frm시트관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  열려진 통합문서 목록과 선택한 통합문서에 포함된 시트의 목록을 표시
'            시트목록의 순서(정렬), 이름 변경 등의 관리 작업이 가능
'------------------------------------------------------------------------------------------
Option Explicit
Option Base 1

'----------------------------------------------------------
'  모든 시트를 표시
'----------------------------------------------------------
Private Sub cmdAllVisible_Click()
   Dim s As Long
   
   On Error Resume Next
   With Workbooks(Me.cboWorkbook.Text)
      For s = 1 To .Sheets.Count
         .Sheets(s).Visible = True
      Next
   End With
   
   Call cboWorkbook_Change
   On Error GoTo 0
End Sub

'----------------------------------------------------------
'  선택한 시트 숨기기
'----------------------------------------------------------
Private Sub cmdVisible_Click()
   Dim s As Long
   On Error Resume Next
   With Me.ListBox1
      If .ListIndex < 0 Then Exit Sub
      s = .ListIndex
      Workbooks(Me.cboWorkbook.Text).Sheets(.List(s, 0)).Visible = IIf(.List(s, 1) = "", False, True)
      
      Call cboWorkbook_Change
      .ListIndex = s
   End With
   On Error GoTo 0
End Sub

'--// 폼이 다시 선택될 때 마다 콤보 상자에
'     현재 열려진 통합문서 이름을 추가하고
Private Sub UserForm_Activate()
   Dim i As Integer
   Me.cboWorkbook.Clear
  For i = 1 To Workbooks.Count
    Me.cboWorkbook.AddItem Workbooks(i).Name
  Next
  '--// 가장 마지막 파일이 선택되도록
  Me.cboWorkbook = Workbooks(Workbooks.Count).Name
End Sub

'----------------------------------------------------------
'  지정한 통합문서의 시트 목록을 콤보 상자에 표시
'----------------------------------------------------------
Private Sub cboWorkbook_Change()
   Dim shtK As Worksheet
   
On Error Resume Next
   If Me.cboWorkbook.Text <> "" Then
      Workbooks(Me.cboWorkbook.Text).Activate
   End If

   With Me.ListBox1
      .Clear
      .ColumnCount = 2
      .ColumnWidths = "160 pt;20 pt"
      For Each shtK In Workbooks(Me.cboWorkbook.Text).Sheets
         .AddItem shtK.Name
         .List(.ListCount - 1, 1) = IIf(shtK.Visible, "", "Hidden")
      Next
      .ListIndex = IIf(.ListCount > 0, .ListCount - 1, -1)
   End With
End Sub


Private Sub cmd목록만들기_Click()
   Dim rngK As Range
   Dim i As Integer
On Error Resume Next
   Set rngK = Application.InputBox _
    ("목록을 출력할 시작 셀을 선택해 주세요.", "셀 지정", Type:=8)
   If rngK Is Nothing Then Exit Sub
   For i = 0 To Me.ListBox1.ListCount - 1
      rngK.Offset(i, 0) = Me.ListBox1.List(i)
   Next
End Sub

Private Sub cmd삭제_Click()
   Dim k As Integer
   
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "선택한 시트가 없습니다."
      Exit Sub
   End If
   If Workbooks(Me.cboWorkbook.Text).Sheets.Count = 1 Then
      MsgBox "총 1개 시트이므로 삭제할 수 없습니다."
      Exit Sub
   End If
 '--// 해당 시트 삭제
   Workbooks(Me.cboWorkbook.Text).Sheets(Me.ListBox1.Text).Delete
   Call cboWorkbook_Change
   If k >= Me.ListBox1.ListCount Then k = k - 1
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd삽입_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "선택한 시트가 없습니다."
      Exit Sub
   End If
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets.Add before:=.Sheets(Me.ListBox1.List(k))
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd아래로_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "선택한 시트가 없습니다."
      Exit Sub
   End If
   If k = Me.ListBox1.ListCount - 1 Then
      MsgBox "가장 끝 시트입니다.", , "시트이동"
      Exit Sub
   End If
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets(k + 1).Move after:=.Sheets(k + 1)
'      .Sheets(Me.ListBox1.Text).Move after:=.Sheets(Me.ListBox1.List(k + 1))
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k + 1
End Sub

Private Sub cmd위로_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "선택한 시트가 없습니다."
      Exit Sub
   End If
   If k = 0 Then
      MsgBox "첫번째 시트입니다.", , "시트이동"
      Exit Sub
   End If
   With Workbooks(Me.cboWorkbook.Text)
      .Activate
      .Sheets(k).Move before:=.Sheets(k)
'      .Sheets(Me.ListBox1.Text).Move before:=.Sheets(Me.ListBox1.List(k - 1))
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k - 1
End Sub

Private Sub cmd이름변경_Click()
   Dim k As Integer
   Dim strName As String
   If Me.ListBox1.ListIndex < 0 Then Exit Sub
   
On Error Resume Next
   strName = InputBox("현재 시트이름 [" & Me.ListBox1.Text & "]을 변경할 새 이름을 입력하세요.", "새 이름")
   If strName = "" Then Exit Sub
   k = Me.ListBox1.ListIndex
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets(Me.ListBox1.Text).Name = strName
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd정렬_Click()
   Dim wkBook As Workbook
   Set wkBook = Workbooks(Me.cboWorkbook.Text)
  If Me.opt오름차순 Then
      Call sb오름차순정렬(wkBook)
   Else
      Call sb내림차순정렬(wkBook)
   End If
   Call cboWorkbook_Change
End Sub


'----------------------------------------------------------
'  지정한 통합문서의 시트들을 오름차순으로 정렬
'----------------------------------------------------------
Sub sb오름차순정렬(wkBook As Workbook)
  Dim shtTemp As Worksheet
  Dim i As Integer, k As Integer
'--// wkSheet 변수는 시트이름을 저장할 배열 변수로
'      크기는 작업할 시트 개수를 계산한 후 ReDim 문으로
'      프로시저 안에서 지정하기위해 괄호만 사용하여
'      동적 배열 변수로 선언
  Dim wkSheet() As String, temp As String
  
  ReDim wkSheet(wkBook.Sheets.Count)
  '선택한 파일의 모든 시트의 이름을 기억 시켜둠
  For i = 1 To wkBook.Sheets.Count
    wkSheet(i) = wkBook.Sheets(i).Name
  Next i
  
 ' 시트 이름을 오름차순 정렬함
   For i = 1 To wkBook.Sheets.Count - 1
     For k = i + 1 To wkBook.Sheets.Count
       If wkSheet(i) > wkSheet(k) Then
         temp = wkSheet(i)
         wkSheet(i) = wkSheet(k)
         wkSheet(k) = temp
       End If
     Next k
   Next i
  '--// 정렬된 내용으로 실제 통합문서의 시트 순서를 변경함
  For i = 1 To wkBook.Sheets.Count
    wkBook.Sheets(wkSheet(i)).Move before:=wkBook.Sheets(i)
  Next
End Sub

'----------------------------------------------------------
'  지정한 통합문서의 시트들을 내림차순으로 정렬
'----------------------------------------------------------
Sub sb내림차순정렬(wkBook As Workbook)
  Dim shtTemp As Worksheet
  Dim i As Integer, k As Integer
  Dim wkSheet() As String, temp As String
  
  ReDim wkSheet(wkBook.Sheets.Count)
  '선택한 파일의 모든 시트의 이름을 기억 시켜둠
  For i = 1 To wkBook.Sheets.Count
    wkSheet(i) = wkBook.Sheets(i).Name
  Next i
  
 ' 시트 이름을 내림차순 정렬함
   For i = 1 To wkBook.Sheets.Count - 1
     For k = i + 1 To wkBook.Sheets.Count
       If wkSheet(i) < wkSheet(k) Then
         temp = wkSheet(i)
         wkSheet(i) = wkSheet(k)
         wkSheet(k) = temp
       End If
     Next k
   Next i
  
  For i = 1 To wkBook.Sheets.Count
    wkBook.Sheets(wkSheet(i)).Move before:=wkBook.Sheets(i)
  Next
End Sub



