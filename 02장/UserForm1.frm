VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�׷캰 ��Ʈ �и� //��� ���ǻ�"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------
'  ��   �� :  �۾������� ������ ���뿡�� ������ ����
'               �������� �׷캰�� ��Ʈ �и��ϴ� �۾�
'  �ۼ��� :  ��� ���� ��ũ�ο� VBA
'-----------------------------------------------------------------
Option Explicit
Dim rng��ü���� As Range

Private Sub cmd�ݱ�_Click()
   Unload Me
End Sub

Private Sub cmd����_Click()
   Dim rng���ؿ� As Range, rngK As Range
   Dim col�׷� As New Collection
   Dim varK As Variant
   Dim i As Long, int���ؿ� As Long
   
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   int���ؿ� = Me.ComboBox1.ListIndex + 1
   
   On Error Resume Next
   For Each rngK In rng��ü����.Columns(int���ؿ�).Cells
      If rngK <> Me.ComboBox1 Then
         If TypeName(rngK.Value) = "String" Then
            col�׷�.Add Item:=rngK, Key:=rngK
         Else
            col�׷�.Add Item:=Trim(rngK.Text), Key:=Trim(rngK.Text)
         End If
      End If
   Next
   
   For Each varK In col�׷�
      If Sheets(varK.Value).Name <> "" Then
         Sheets(varK.Value).Delete
      End If
      If TypeName(rng��ü����.Cells(2, int���ؿ�).Value) = "Date" Then
         rng��ü����.AutoFilter Field:=int���ؿ�, Operator:= _
            xlFilterValues, Criteria2:=Array(2, varK)
      Else
      rng��ü����.AutoFilter Field:=int���ؿ�, Criteria1:=varK
      End If
      rng��ü����.Copy
      
      Sheets.Add After:=Sheets(Sheets.Count)
      ActiveSheet.Name = varK
      ActiveSheet.Paste
      Selection.Columns.AutoFit
    Next
   On Error GoTo 0
    
   rng��ü����.AutoFilter
   Application.CutCopyMode = False
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, "�Ϸ�"
   Unload Me
End Sub


Private Sub RefEdit1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
   Dim rngK As Range
   If Trim(Me.RefEdit1.Text) = "" Then Exit Sub
   Set rng��ü���� = Range(Me.RefEdit1)
   If rng��ü����.Areas.Count > 1 Then
      MsgBox "�� ������ ���ӵ� ���� �����̿��� �մϴ�.", vbCritical, "���� ���� ����"
      Exit Sub
   End If
   
   Me.ComboBox1.Clear
   For Each rngK In rng��ü����.Rows(1).Cells
      Me.ComboBox1.AddItem rngK
   Next
End Sub
