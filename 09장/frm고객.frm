VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm�� 
   Caption         =   "������ ��ȸ"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
   OleObjectBlob   =   "frm��.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbo�����׸�_Click()
   Dim keyCol As Integer
   Dim rngT As Range
   Set rngT = Range("nm��_���").CurrentRegion
   keyCol = Application.Match(Me.cbo�����׸�, rngT.Rows(1), 0) + rngT.Column - 1
   '--// ���� �޼���� 2007 ���� ����� : 2003 ���Ͽ��� ���� �߻�
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

Private Sub cmd�˻�_Click()
   Dim rngT As Range
   
   '--// <�����> ��Ʈ�� �� �� ����  'txt�˻��̸�' ������ �����ϴ� �ڷḸ
   '      ������͸� �̿��Ͽ�  <�۾���Ʈ>�� ���
   Range("nm��_����").Cells(2, 1) = "*" & Me.txt�˻��̸� & "*"
   Range("tbl������[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm��_����"), _
        CopyToRange:=Range("nm��_���"), Unique:=False
        
   '--// <�۾���Ʈ>�� ��µ� ������ lst����� ��� ���ڿ� ǥ��
   Set rngT = Range("nm��_���").CurrentRegion
   
   With Me.lst�����
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
   
   '--// ���� �׸����� �ڷ� ����
   Call cbo�����׸�_Click
End Sub

Private Sub cmd����_Click()
   Dim iOK As Integer, R As Long
   
   iOK = MsgBox("���� ��ȸ ���� �ڷḦ �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "����Ȯ��")
   If iOK = vbYes Then
      R = Application.Match(Me.txt���ڵ�, Range("tbl������[���ڵ�]"), 0)
      If R > 0 Then
         R = R + Range("tbl������[[#Headers],[���ڵ�]]").Row
         Sheets("�����").Rows(R).Delete Shift:=xlUp
         MsgBox "������ �Ϸ�Ǿ����ϴ�.", vbInformation
      End If
      
      '--// ������ ��� ���� ���� �ٽ� ǥ��
      Call cmd�˻�_Click
   End If
End Sub

Private Sub cmd����_Click()
   Call sb��Ʈ�����(False)
End Sub

Private Sub cmd�ű�_Click()
   Call sb��Ʈ�����(False)
   Call sb��Ʈ�ѳ������
   
   '--// �ڵ带 ���� �ο�. �ڵ带 �ڵ� �ο��ϱ� ���� <�۾���Ʈ> A2 ���� �̸�
   '     �迭������ �̿��Ͽ� ���ڵ� �� �ִ� ���ڸ� ����ϰ� ����
   Me.txt���ڵ� = "S" & Format(Range("nmMax�ڵ�") + 1, "00000")
   Me.txt����.SetFocus
   Call sb��ưǥ��(False)
End Sub

Private Sub cmd����_Click()
   Dim R As Long
   
   If Me.txt���ڵ� = "" Or Me.txt���� = "" Then
      MsgBox "���ڵ�� ������ �Է��ϼ���.", vbCritical
      Exit Sub
   End If
   '--// <�����> ��Ʈ���� txt���ڵ� ��Ʈ���� ���ڵ尡 ���° �࿡ ��ġ�ϴ���
   '    Ȯ��. ��ã�� ��� �ű� ����ϱ� ���� ���� �ڷ��� ���� ������ �� ���� ���� ��ȯ
   If Application.CountIf(Range("tbl������[���ڵ�]"), Me.txt���ڵ�) = 0 Then
      R = Range("tbl������[���ڵ�]").End(xlDown).Row + 1
   Else
      R = Application.Match(Me.txt���ڵ�, Range("tbl������[���ڵ�]"), 0)
      R = R + Range("tbl������[[#Headers],[���ڵ�]]").Row
   End If
   
   With Sheets("�����")
      .Cells(R, 1) = Me.txt���ڵ�
      .Cells(R, 2) = Me.txt����
      .Cells(R, 3) = Me.txt�Ҽ�
      .Cells(R, 4) = Me.txt����ó
      .Cells(R, 5) = Me.txt�ּ�
   
   MsgBox "����Ǿ����ϴ�.", vbInformation
   
   '--// ��ϵ� ������ ��� ���ڿ� �ݿ��ǵ��� <�˻�> ��ư�� Ŭ���� ��ó�� ����
   Call cmd�˻�_Click
   '--// �ֱ� ����/����� ����� ǥ�õǵ��� ��� ���� ����
   Me.lst�����.Text = .Cells(R, 1)
   End With
End Sub

Private Sub cmd���_Click()
   Call lst�����_Click
End Sub

Private Sub lst�����_Click()
   With Me.lst�����
      If .ListIndex >= 0 Then
         Me.txt���ڵ� = .Column(0)
         Me.txt���� = .Column(1)
         Me.txt�Ҽ� = .Column(2)
         Me.txt����ó = .Column(3)
         Me.txt�ּ� = .Column(4)
      Else
         Call sb��Ʈ�ѳ������
      End If
   End With
   
   Call sb��Ʈ�����(True)
End Sub


Private Sub txt�˻��̸�_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = 13 Then
      Call cmd�˻�_Click
      With Me.txt�˻��̸�
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
      End With
      KeyCode = 0       '--// ����Ű�� ��ȿ ���Ѽ� �ٸ� ��Ʈ�ѷ� Ŀ���� �̵����� �ʵ��� ��
   End If
End Sub


Private Sub UserForm_Initialize()
   With Me.cbo�����׸�
      .Clear
      .AddItem "���ڵ�"
      .AddItem "����"
      .AddItem "�Ҽ�"
      .AddItem "����ó"
      .AddItem "�ּ�"
      .Text = "���ڵ�"
   End With
   
   Call cmd�˻�_Click
End Sub

'-------------------------------------------------------------
' ������ �ű� ��� ���¿����� �ؽ�Ʈ ���ڰ� ����� �� �ֵ��� ó��
' ��Ʈ�� Ư��ȿ��(SpecialEffect)�� ��Ȳ�� ���� �޶����� ��
'-------------------------------------------------------------
Sub sb��Ʈ�����(bLock As Boolean)
   Dim ctrNM
   Dim i As Integer
   ctrNM = Array("txt����", "txt�Ҽ�", "txt����ó", "txt�ּ�")
   
   For i = LBound(ctrNM) To UBound(ctrNM)
      Me.Controls(ctrNM(i)).Locked = bLock
      Me.Controls(ctrNM(i)).SpecialEffect = IIf(bLock, 3, 2)
   Next
   
   Call sb��ưǥ��(bLock)
End Sub

Sub sb��Ʈ�ѳ������()
   Me.txt���ڵ� = ""
   Me.txt���� = ""
   Me.txt�Ҽ� = ""
   Me.txt����ó = ""
   Me.txt�ּ� = ""
End Sub

'-------------------------------------------------------------
' ������ �ű� ��� ���¿����� ����/��� ��ư�� ǥ�õǰ�
' �� �̿ܿ��� ������ϱ� ���� ó��
'-------------------------------------------------------------
Sub sb��ưǥ��(bShow As Boolean)
   Me.cmd����.Visible = bShow
   Me.cmd����.Visible = bShow
   Me.cmd�ű�.Visible = bShow
   Me.cmd����.Visible = Not bShow
   Me.cmd���.Visible = Not bShow
End Sub

