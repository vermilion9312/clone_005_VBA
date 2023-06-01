VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm�׷�и� 
   Caption         =   "�׷캰 ��Ʈ �и� //��� ���ǻ�"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   OleObjectBlob   =   "frm�׷�и�.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "frm�׷�и�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ǥ ������ ����Ʈ�� �����ϸ�, ù ���� �ʵ������ �ν��� ��
'            �ش� �ʵ忡 ���� �ڷ���� �и��Ͽ� ������ ��Ʈ�� ����
'------------------------------------------------------------------------------------------

Option Explicit
Dim rngTable As Range

'---------------------------------------------------------------
' �� ����� ���� ���� ������ �ִ� ��� ���� �� ������
' ������ �ڵ� ����
'---------------------------------------------------------------
Private Sub UserForm_Initialize()
   On Error Resume Next
   lblMsg.Caption = "�۾� ������ �и� ���� ������ �� <����>�� Ŭ���ϼ���."
   If ActiveCell <> "" Then
      Me.RefEdit1.Text = "'" & ActiveSheet.Name & "'!" & ActiveCell.CurrentRegion.Address
   End If
   On Error GoTo 0
End Sub

Private Sub RefEdit1_Change()
   lblMsg.Caption = "�۾� ������ �и� ���� ������ �� <����>�� Ŭ���ϼ���."
End Sub

Private Sub cboCol_Change()
   lblMsg.Caption = "�۾� ������ �и� ���� ������ �� <����>�� Ŭ���ϼ���."
End Sub

Private Sub cmd�ݱ�_Click()
   Unload Me
End Sub

'---------------------------------------------------------------
' �۾� ������ ������ �� ��Ʈ���� �������� ��
' cboCol ��Ʈ���� ����� �۾� ���� ù ���� ��������  �ٽ� ǥ��
'---------------------------------------------------------------
Private Sub RefEdit1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
   Dim rngK As Range
   If Trim(Me.RefEdit1.Text) = "" Then Exit Sub
   Set rngTable = Range(Me.RefEdit1)
   If rngTable.Areas.Count > 1 Then
      MsgBox "�� ������ ���ӵ� ���� �����̿��� �մϴ�.", vbCritical, "���� ���� ����"
      Exit Sub
   End If
   
   Me.cboCol.Clear
   For Each rngK In rngTable.Rows(1).Cells
      Me.cboCol.AddItem rngK
   Next
End Sub

'---------------------------------------------------------------
' �۾� ����
'---------------------------------------------------------------
Private Sub cmd����_Click()
   Dim GroupList  As New Collection
   Dim varK As Variant
   Dim R As Long, ColNo As Long
   Dim rngK As Range
   Dim wkB As Workbook, sh As Worksheet
   
   If Me.RefEdit1.Text = "" Or Me.cboCol.Text = "" Then Exit Sub
   
   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
   
   ColNo = Me.cboCol.ListIndex + 1    '--// ���� ��ġ��ȣ
   
   '--// GroupList ������ ������ ��(ColNo)�� �ڷ���� �ߺ����� �ϳ��� ����(�迭 ���� ���°� ��)
   On Error Resume Next
   '--// ������ ������ ù ���� �ʵ���̱⶧���� �����ϰ� �ι�° ������ ����
   For R = 2 To rngTable.Columns(ColNo).Cells.Count
      Set rngK = rngTable.Columns(ColNo).Cells(R, 1)
      If TypeName(rngK.Value) = "Date" Then
         '--// �������� ��¥�� ��쿡�� �� ���� ������� ������ �ߺ� üũ
         GroupList.Add Item:=rngK.Value, Key:="D" & rngK.Value
      Else
         '--// �ڵ� ���Ϳ��� ��¥�� ������ ������ ��� ������ ����� �ؽ�Ʈ�� �ν�
         GroupList.Add Item:=rngK.Text, Key:=rngK.Text
      End If
   Next
   On Error GoTo 0
   
On Error GoTo End_Rtn
   If GroupList.Count = 0 Then Exit Sub
   Me.MousePointer = fmMousePointerHourGlass '--// ������ ���콺 ������ ����� �𷡽ð�� ����
   
   '--// �� ��ũ�� �߰�
   Set wkB = Workbooks.Add
   
   '--// ����� GroupList�� �������� �ڵ��������� ���� ����� �����Ͽ� �� ��ũ��Ʈ�� �ٿ��ֱ��Ͽ� ����
   For Each varK In GroupList
      '--// ���� �޽����� ǥ��
      lblMsg.Caption = "��" & varK & "���� ���� �и� �۾��� ���� ���Դϴ�. ��ø� ��ٸ�����."
      Me.Repaint
      
      '--// �ʵ��� ������ ������ ��¥���� ���� ��/���� ���еǰ� ����
      If TypeName(rngTable.Cells(2, ColNo).Value) = "Date" Then
         rngTable.AutoFilter Field:=ColNo, Operator:=xlFilterValues, Criteria2:=Array(2, varK)
      Else
         rngTable.AutoFilter Field:=ColNo, Criteria1:=varK
      End If
      rngTable.Copy
      
      wkB.Sheets.Add After:=Sheets(Sheets.Count)
      ActiveSheet.Name = varK
      ActiveSheet.Paste
      Selection.Columns.AutoFit
   Next
   '--// ���ʿ��� ��Ʈ ����
   For Each sh In wkB.Sheets
      If sh.UsedRange.Address = "$A$1" Then sh.Delete
   Next
    
   '--// �ڵ����Ͱ� �Ǿ��ִ� ���¿��� ���� ���� ��� �����
   If rngTable.Parent.AutoFilter.FilterMode Then rngTable.Parent.AutoFilter.ShowAllData
   wkB.Activate
   
End_Rtn:
   Application.CutCopyMode = False
   Me.MousePointer = fmMousePointerDefault  '--// ������ ���콺 ������ ����� �⺻���� ����
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   If Err.Number = 0 Then
      lblMsg.Caption = "�۾��� �Ϸ�Ǿ����ϴ�."
      MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, "�Ϸ�"
   Else
      lblMsg.Caption = "�۾� �� ������ ���� ������ �߻��߽��ϴ�."
      MsgBox "�۾� �� ������ ���� ������ �߻��߽��ϴ�." & vbCrLf & _
               Err.Description, vbCritical, "����"
   End If
End Sub
