VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm���� 
   Caption         =   "������ȸ �� ����"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "frm����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
   Dim i As Long
   With Me.lst������
      For i = 0 To .ListCount - 1
         .Selected(i) = chkAll.Value
      Next
   End With
End Sub

Private Sub cmd�˻�_Click()
   Dim rngT As Range
   
   With Range("nm����_����")
      .Cells(2, 1) = "*" & Me.txt�˻��̸� & "*"
      .Cells(2, 2) = IIf(Me.txtFrom = "", "", ">=") & Me.txtFrom
      .Cells(2, 3) = IIf(Me.txtTo = "", "", "<=") & Me.txtTo
   End With
   
   Range("tbl��������[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm����_����"), _
        CopyToRange:=Range("nm����_���"), Unique:=False
        
   Set rngT = Range("nm����_���").CurrentRegion
   
   With Me.lst���¸��
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


Private Sub cmd������_Click()
   Dim R As Long, Cnt As Long, i As Long
   Dim rngT As Range
   
   With Sheets("���-���º�")
      .Range("C3") = Me.txt���¸� & " (" & Me.txt�����ڵ� & ")"
      .Range("C4") = Format(Me.txt����, "yy-mm-dd(aaa)")
      .Range("C5") = Me.txt���
      .Range("G3") = Me.lst���¸��.Column(5)
      Set rngT = .Range("A8")
      '--// 8�� ���� ����� ���� ����
      .Range(rngT.Offset(1, 0), rngT.SpecialCells(xlLastCell)).Clear
   End With
      
   With Me.lst������
      rngT.EntireRow.Copy
      With rngT.Offset(1, 0).Resize(.ListCount - 1).EntireRow
            .PasteSpecial Paste:=xlPasteFormats
            .PasteSpecial Paste:=xlPasteFormulas
      End With
      Application.CutCopyMode = False
      
      For i = 0 To .ListCount - 1
         rngT.Offset(i, 0) = i + 1
         rngT.Offset(i, 1) = .List(i, 0)  '--// ���ڵ�
         rngT.Offset(i, 3) = .List(i, 1)  '--// ����
         rngT.Offset(i, 6) = Format(.List(i, 2), "yy-mm-dd(aaa)") '--//������
         rngT.Offset(i, 7) = .List(i, 3)  '--// ���
         rngT.Offset(i, 8) = .List(i, 4)  '--// ��������
      Next
      '--// �μ� ���� �ٽ� ����
      Sheets("���-���º�").PageSetup.PrintArea = "$A$1:$I$" & (rngT.Row + .ListCount)
      Sheets("���-���º�").PrintOut preview:=True
   End With
End Sub

Private Sub cmd����_Click()
   Dim i As Long, R As Long, iOK As Integer
   Dim strKey
  
   iOK = MsgBox("�����Ͻ� �ڷ����  �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "����Ȯ��")
   If iOK = vbYes Then
      With Me.lst������
         For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
               strKey = Me.txt�����ڵ� & .List(i, 0)
               If Application.CountIf(Range("tbl������Ȳ[Key]"), strKey) > 0 Then
                  R = Application.Match(strKey, Range("tbl������Ȳ[Key]").EntireColumn, 0)
                  Sheets("������Ȳ").Rows(R).Delete Shift:=xlUp
               End If
            End If
         Next
      End With
      
      MsgBox "������ �Ϸ�Ǿ����ϴ�.", vbInformation
      Call sb���������
   End If
End Sub

Private Sub cmd�ű�_Click()
   Dim i As Long, R As Long
   Dim strKey
   Application.DisplayAlerts = False
   
   With Me.lst�����
      For i = 0 To .ListCount - 1
         If .Selected(i) Then
            strKey = Me.txt�����ڵ� & .List(i, 0)
            If Application.CountIf(Range("tbl������Ȳ[Key]"), strKey) = 0 Then
               If Range("tbl������Ȳ[Key]").Rows.Count = 1 Then
                  R = Range("tbl������Ȳ[Key]").Row + 1
               Else
                  R = Range("tbl������Ȳ[Key]").End(xlDown).Row + 1
               End If
               
               With Sheets("������Ȳ")
                  .Cells(R, 2) = Me.txt�����ڵ�
                  .Cells(R, 3) = Me.lst�����.List(i, 0)
                  .Cells(R, 5) = Date
               End With
            End If
         End If
      Next
   End With
   
   Call sb���������
   Application.DisplayAlerts = True

End Sub


Private Sub lst���¸��_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Call sb���������
   Call sb��������
   Me.MultiPage1.Value = 1
End Sub

Sub sb���������()
   Dim rngT As Range
   
   '--//  [������ȸ] ���������� ������ �������� ǥ��
   With Me.lst���¸��
      Me.txt�����ڵ� = .Column(0)
      Me.txt���¸� = .Column(2)
      Me.txt���� = .Column(1)
      Me.txt��� = .Column(3)
   End With
   
   '--// [������Ȳ] ��Ʈ�� ���� �� ������ ���� �ڵ��� ������ ������ͷ� ����
   Range("nm����_����").Cells(2, 1) = Me.txt�����ڵ�
   Range("tbl������Ȳ[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("nm����_����"), _
        CopyToRange:=Range("nm����_���"), Unique:=False
        
   '--// ������� ��� ������ ������ ��� ���ڿ� ǥ��
   Set rngT = Range("nm����_���").CurrentRegion
   
   With Me.lst������
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

Sub sb��������()
   Dim rngT As Range
   '--// [�����] ��Ʈ�� �� ����� '�����' ��� ���ڿ� ���
   Set rngT = Range("tbl������[#All]")
   
   With Me.lst�����
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
   If Me.MultiPage1.Value = 1 And Me.txt�����ڵ� = "" Then
      Call sb���������
      Call sb��������
   End If
End Sub

Private Sub UserForm_Initialize()
   Me.MultiPage1.Value = 0
   Call cmd�˻�_Click
   chkAll = False
End Sub
