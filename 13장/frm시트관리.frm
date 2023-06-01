VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm��Ʈ���� 
   Caption         =   "��Ʈ���� "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   OleObjectBlob   =   "frm��Ʈ����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm��Ʈ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ������ ���չ��� ��ϰ� ������ ���չ����� ���Ե� ��Ʈ�� ����� ǥ��
'            ��Ʈ����� ����(����), �̸� ���� ���� ���� �۾��� ����
'------------------------------------------------------------------------------------------
Option Explicit
Option Base 1

'----------------------------------------------------------
'  ��� ��Ʈ�� ǥ��
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
'  ������ ��Ʈ �����
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

'--// ���� �ٽ� ���õ� �� ���� �޺� ���ڿ�
'     ���� ������ ���չ��� �̸��� �߰��ϰ�
Private Sub UserForm_Activate()
   Dim i As Integer
   Me.cboWorkbook.Clear
  For i = 1 To Workbooks.Count
    Me.cboWorkbook.AddItem Workbooks(i).Name
  Next
  '--// ���� ������ ������ ���õǵ���
  Me.cboWorkbook = Workbooks(Workbooks.Count).Name
End Sub

'----------------------------------------------------------
'  ������ ���չ����� ��Ʈ ����� �޺� ���ڿ� ǥ��
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


Private Sub cmd��ϸ����_Click()
   Dim rngK As Range
   Dim i As Integer
On Error Resume Next
   Set rngK = Application.InputBox _
    ("����� ����� ���� ���� ������ �ּ���.", "�� ����", Type:=8)
   If rngK Is Nothing Then Exit Sub
   For i = 0 To Me.ListBox1.ListCount - 1
      rngK.Offset(i, 0) = Me.ListBox1.List(i)
   Next
End Sub

Private Sub cmd����_Click()
   Dim k As Integer
   
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "������ ��Ʈ�� �����ϴ�."
      Exit Sub
   End If
   If Workbooks(Me.cboWorkbook.Text).Sheets.Count = 1 Then
      MsgBox "�� 1�� ��Ʈ�̹Ƿ� ������ �� �����ϴ�."
      Exit Sub
   End If
 '--// �ش� ��Ʈ ����
   Workbooks(Me.cboWorkbook.Text).Sheets(Me.ListBox1.Text).Delete
   Call cboWorkbook_Change
   If k >= Me.ListBox1.ListCount Then k = k - 1
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd����_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "������ ��Ʈ�� �����ϴ�."
      Exit Sub
   End If
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets.Add before:=.Sheets(Me.ListBox1.List(k))
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd�Ʒ���_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "������ ��Ʈ�� �����ϴ�."
      Exit Sub
   End If
   If k = Me.ListBox1.ListCount - 1 Then
      MsgBox "���� �� ��Ʈ�Դϴ�.", , "��Ʈ�̵�"
      Exit Sub
   End If
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets(k + 1).Move after:=.Sheets(k + 1)
'      .Sheets(Me.ListBox1.Text).Move after:=.Sheets(Me.ListBox1.List(k + 1))
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k + 1
End Sub

Private Sub cmd����_Click()
   Dim k As Integer
   k = Me.ListBox1.ListIndex
   If k < 0 Then
      MsgBox "������ ��Ʈ�� �����ϴ�."
      Exit Sub
   End If
   If k = 0 Then
      MsgBox "ù��° ��Ʈ�Դϴ�.", , "��Ʈ�̵�"
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

Private Sub cmd�̸�����_Click()
   Dim k As Integer
   Dim strName As String
   If Me.ListBox1.ListIndex < 0 Then Exit Sub
   
On Error Resume Next
   strName = InputBox("���� ��Ʈ�̸� [" & Me.ListBox1.Text & "]�� ������ �� �̸��� �Է��ϼ���.", "�� �̸�")
   If strName = "" Then Exit Sub
   k = Me.ListBox1.ListIndex
   With Workbooks(Me.cboWorkbook.Text)
      .Sheets(Me.ListBox1.Text).Name = strName
   End With
   Call cboWorkbook_Change
   Me.ListBox1.ListIndex = k
End Sub

Private Sub cmd����_Click()
   Dim wkBook As Workbook
   Set wkBook = Workbooks(Me.cboWorkbook.Text)
  If Me.opt�������� Then
      Call sb������������(wkBook)
   Else
      Call sb������������(wkBook)
   End If
   Call cboWorkbook_Change
End Sub


'----------------------------------------------------------
'  ������ ���չ����� ��Ʈ���� ������������ ����
'----------------------------------------------------------
Sub sb������������(wkBook As Workbook)
  Dim shtTemp As Worksheet
  Dim i As Integer, k As Integer
'--// wkSheet ������ ��Ʈ�̸��� ������ �迭 ������
'      ũ��� �۾��� ��Ʈ ������ ����� �� ReDim ������
'      ���ν��� �ȿ��� �����ϱ����� ��ȣ�� ����Ͽ�
'      ���� �迭 ������ ����
  Dim wkSheet() As String, temp As String
  
  ReDim wkSheet(wkBook.Sheets.Count)
  '������ ������ ��� ��Ʈ�� �̸��� ��� ���ѵ�
  For i = 1 To wkBook.Sheets.Count
    wkSheet(i) = wkBook.Sheets(i).Name
  Next i
  
 ' ��Ʈ �̸��� �������� ������
   For i = 1 To wkBook.Sheets.Count - 1
     For k = i + 1 To wkBook.Sheets.Count
       If wkSheet(i) > wkSheet(k) Then
         temp = wkSheet(i)
         wkSheet(i) = wkSheet(k)
         wkSheet(k) = temp
       End If
     Next k
   Next i
  '--// ���ĵ� �������� ���� ���չ����� ��Ʈ ������ ������
  For i = 1 To wkBook.Sheets.Count
    wkBook.Sheets(wkSheet(i)).Move before:=wkBook.Sheets(i)
  Next
End Sub

'----------------------------------------------------------
'  ������ ���չ����� ��Ʈ���� ������������ ����
'----------------------------------------------------------
Sub sb������������(wkBook As Workbook)
  Dim shtTemp As Worksheet
  Dim i As Integer, k As Integer
  Dim wkSheet() As String, temp As String
  
  ReDim wkSheet(wkBook.Sheets.Count)
  '������ ������ ��� ��Ʈ�� �̸��� ��� ���ѵ�
  For i = 1 To wkBook.Sheets.Count
    wkSheet(i) = wkBook.Sheets(i).Name
  Next i
  
 ' ��Ʈ �̸��� �������� ������
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



