VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "그룹별 시트 분리 //길벗 출판사"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------
'  기   능 :  작업범위로 지정한 내용에서 지정한 열을
'               기준으로 그룹별로 시트 분리하는 작업
'  작성자 :  길벗 엑셀 매크로와 VBA
'-----------------------------------------------------------------
Option Explicit
Dim rng전체범위 As Range

Private Sub cmd닫기_Click()
   Unload Me
End Sub

Private Sub cmd실행_Click()
   Dim rng기준열 As Range, rngK As Range
   Dim col그룹 As New Collection
   Dim varK As Variant
   Dim i As Long, int기준열 As Long
   
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   int기준열 = Me.ComboBox1.ListIndex + 1
   
   On Error Resume Next
   For Each rngK In rng전체범위.Columns(int기준열).Cells
      If rngK <> Me.ComboBox1 Then
         If TypeName(rngK.Value) = "String" Then
            col그룹.Add Item:=rngK, Key:=rngK
         Else
            col그룹.Add Item:=Trim(rngK.Text), Key:=Trim(rngK.Text)
         End If
      End If
   Next
   
   For Each varK In col그룹
      If Sheets(varK.Value).Name <> "" Then
         Sheets(varK.Value).Delete
      End If
      If TypeName(rng전체범위.Cells(2, int기준열).Value) = "Date" Then
         rng전체범위.AutoFilter Field:=int기준열, Operator:= _
            xlFilterValues, Criteria2:=Array(2, varK)
      Else
      rng전체범위.AutoFilter Field:=int기준열, Criteria1:=varK
      End If
      rng전체범위.Copy
      
      Sheets.Add After:=Sheets(Sheets.Count)
      ActiveSheet.Name = varK
      ActiveSheet.Paste
      Selection.Columns.AutoFit
    Next
   On Error GoTo 0
    
   rng전체범위.AutoFilter
   Application.CutCopyMode = False
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   MsgBox "작업이 완료되었습니다.", vbInformation, "완료"
   Unload Me
End Sub


Private Sub RefEdit1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
   Dim rngK As Range
   If Trim(Me.RefEdit1.Text) = "" Then Exit Sub
   Set rng전체범위 = Range(Me.RefEdit1)
   If rng전체범위.Areas.Count > 1 Then
      MsgBox "셀 영역은 연속된 단일 영역이여야 합니다.", vbCritical, "다중 영역 오류"
      Exit Sub
   End If
   
   Me.ComboBox1.Clear
   For Each rngK In rng전체범위.Rows(1).Cells
      Me.ComboBox1.AddItem rngK
   Next
End Sub
