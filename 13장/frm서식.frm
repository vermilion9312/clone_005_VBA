VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm���� 
   Caption         =   "Ư�� �ܾ� ���ĸ� �����ϱ� // ��� ���ǻ�"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   OleObjectBlob   =   "frm����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ������ Ư�� �ܾ ã�� ������ ������ �� ���
'   �Բ� ����ؾ��� �ڷ� : modPictureSaveAPI , frm����
'------------------------------------------------------------------------------------------

Option Explicit
Option Compare Text  '������ �񱳽� ��ҹ��ڸ� �������� �ʰ� ��

Private Sub cmd�۲ü���_Click()
   Dim wkA As Worksheet
   
   Application.ScreenUpdating = False
   
   If Me.txtã���ܾ�.Text = "" Then
      ThisWorkbook.Sheets("#Color").Range("A1").Value = "������ �۲� ����"
   Else
      ThisWorkbook.Sheets("#Color").Range("A1").Value = Me.txtã���ܾ�
   End If
   
   Set wkA = ActiveSheet   '--// ���� �����ϰ� �ִ� ��Ʈ�� ���
   '--// �۲� ������ �����ϱ� ����, '#Color' ��Ʈ�� A1 ���� ������
   ThisWorkbook.Sheets("#Color").Visible = True
   ThisWorkbook.Sheets("#Color").Select
   Range("A1").Select
   '--// Dialogs(xlDialogFontProperties)�� �̿��ϸ� '�۲� ����' ��ȭ���ڰ� ��Ÿ����
   '     ���� ������ ������ �ڵ����� ������ �۲� ������ ������
   Application.Dialogs(xlDialogFontProperties).Show
   wkA.Select     '--// ó�� �����ߴ� ��Ʈ�� �ٽ� ����
   
   Call �׸���������
   
   Application.ScreenUpdating = True
End Sub

Sub �׸���������()
   Dim sFile As String
   '--// �۲� ������ ����� A1 �� ������ �׸����� ����
   sFile = ThisWorkbook.Path & "\temp.jpg"
   dhSavePic ThisWorkbook.Sheets("#Color").Range("A1"), sFile
    
   '--// �̹��� ��Ʈ�ѿ� ������ �׸��� ǥ��
   Me.Image1.Picture = LoadPicture(sFile)
   Me.Image1.PictureSizeMode = fmPictureSizeModeZoom
   Kill sFile
   ThisWorkbook.Sheets("#Color").Visible = False
End Sub

Private Sub cmd�ݱ�_Click()
   Unload Me
End Sub

Private Sub UserForm_Initialize()
   '--// RGB ���� �������� ���� ��Ʈ ����
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
      .Value = "������ �۲� ����"
      .ColumnWidth = 30
      .RowHeight = 60
   End With

   Call �׸���������
End Sub

Private Sub cmdȮ��_Click()
   Dim rngFind As Range, rngFirst As Range, rngWork As Range, rngK As Range
   Dim iST As Long, iLen As Long, i As Long
   iLen = Len(Me.txtã���ܾ�)
   
   If Me.RefEdit1 = "" Then
      MsgBox "�۾� ������ �����ϼ���.", vbCritical, "�������� ����"
      Exit Sub
   End If
   
  
   Set rngWork = Range(Me.RefEdit1)    '--// �ܾ ã�� ������ rngWork ������ ����
   Set rngK = ThisWorkbook.Sheets("#Color").Range("A1")  '--// ������ ������ ����� ���� rngK ������ ����
   If rngWork.Cells.Count = 1 Then
      '--// ���� ������ �� ���� ��� �ش� ���� ã�� ������ �ִ��� Ȯ��
      If InStr(rngWork, Me.txtã���ܾ�) > 0 Then
         Set rngFirst = rngWork
      End If
   Else
      '--// ���� ������ ���� ���� ��� 'ã�� ���'(Find)�� �̿��� ã��
      Set rngFirst = rngWork.Find(What:=Me.txtã���ܾ�, After:=rngWork.Range("A1"), _
         LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
         SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
   End If
   '--// ù��° ã�� ���� rngFirst ������ ������ �ξ�, Find �˻��� �����
   '   rngFirst ������  ���� ���̵� �� ã�� ����� ������
   Set rngFind = rngFirst
   
   If rngFind Is Nothing Then
      MsgBox "�ش� �ڷḦ ã�� �� �����ϴ�.", vbCritical, "�ڷ� ����"
      Exit Sub
   End If
   
   '--// ������ �˻� ����(rngWork)���� '���� ã��'�� �̿��� �ش� �ܾ �˻��Ͽ�
   '     ������ ����
   i = 0
   Do
      i = i + 1
      iST = InStr(rngFind, Me.txtã���ܾ�)
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
    
    MsgBox i & "���� ã�� ������ �����߽��ϴ�.", vbInformation, "�۾� �Ϸ�"
    
End Sub

