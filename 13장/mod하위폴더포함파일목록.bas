Attribute VB_Name = "mod���������������ϸ��"
'-----------------------------------------------------------------
'  ��  �� : �������� �����Ͽ� ���� ��� �����
'  �ۼ��� :  ��� ���� ��ũ�ο� VBA ���̺�
'-----------------------------------------------------------------

'--// ����� ���� ���� ��Ʈ ��� ǥ���� �� ��Ʈ �̵�

'--// ���������� �����Ͽ� ���� ����Ʈ
Dim rngS As Range
Dim Cnt As Long
Dim subFolder As String

Sub sbFile_List()
  Dim i As Long
  Dim fdNm
  Dim iOk As Integer
  
  iOk = MsgBox("������ ���� ����� �˻��մϴ�." _
            & vbCrLf & "���� ������ �����Ͽ� �˻� �ұ��?" _
            , vbQuestion + vbYesNo, "���� ���")
  ' ����, ������, ���ϸ�, ũ��, ��������
  Set rngS = ThisWorkbook.Sheets(1).Range("A1")
  With rngS
    .CurrentRegion.Clear
    .Offset(0, 0) = "��������"
    .Offset(0, 1) = "�̸�"
    .Offset(0, 2) = "����"
    .Offset(0, 3) = "ũ��(Byte)"
    .Offset(0, 4) = "�ۼ���"
    .CurrentRegion.Interior.Color = vbGreen
  End With
  Cnt = 0
    
  Application.ScreenUpdating = False
  
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = False Then Exit Sub
    
    subFolder = ""
    Call sbFolder_Scan(.SelectedItems(1))
    
    Do While subFolder <> "" And (iOk = vbYes)
      fdNm = Split(Mid(subFolder, 2), ",")
      subFolder = ""
      
      For i = LBound(fdNm) To UBound(fdNm)
         Call sbFolder_Scan(CStr(fdNm(i)))
      Next
    Loop
  End With

'--// ����
  rngS.CurrentRegion.Columns.AutoFit
  With rngS.Parent.Sort
    .SortFields.Clear  '���� ������ �ʱ�ȭ
    .SortFields.Add Key:=rngS, Order:=xlAscending
    .SortFields.Add Key:=rngS.Offset(0, 1), Order:=xlAscending
    .SetRange rngS.CurrentRegion
    .Header = xlYes
    .Apply
  End With
    
'--// ���Ǻ� ����
  Cells.FormatConditions.Delete
  Dim fc As FormatCondition
  
  With Range("A1").CurrentRegion
    Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$C1=""Folder""")
    
    With fc.Interior
      .ThemeColor = xlThemeColorAccent6
      .TintAndShade = 0.5
    End With

  End With
    
'--// ���� ����
  Range("A1").CurrentRegion.AutoFilter

'--// Ʋ����
  Range("A1").Activate
  ActiveWindow.FreezePanes = True
  
  Application.ScreenUpdating = True
End Sub

Sub sbFolder_Scan(strPath As String)
  Dim strFile  'As String
  
  If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
  
  strFile = Dir(strPath, vbDirectory)
  Do While strFile <> ""
     If strFile <> "." And strFile <> ".." Then
        Cnt = Cnt + 1
        rngS.Offset(Cnt, 0) = strPath
        rngS.Offset(Cnt, 1) = strFile
        rngS.Offset(Cnt, 3) = FileLen(strPath & strFile)
        rngS.Offset(Cnt, 4) = FileDateTime(strPath & strFile)
    
        If (GetAttr(strPath & strFile) And vbDirectory) = vbDirectory Then
            rngS.Offset(Cnt, 2) = "Folder"
            
            subFolder = subFolder & "," & strPath & strFile
        Else
          rngS.Offset(Cnt, 2) = "File"
        End If
     End If
     
     strFile = Dir
  Loop
End Sub

