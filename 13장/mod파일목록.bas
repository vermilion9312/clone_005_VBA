Attribute VB_Name = "mod���ϸ��"
Option Explicit
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  ���� ã�� ��ȭ���ڸ� ���� �������� ����� ǥ���� ��
'          ������ ������ ����� ���ϰ� �������� ������ �� ���չ����� ǥ��
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
' ��� :  ���� ã�� ��ȭ���ڸ� ���� �������� ����� ǥ���� ��
'          ������ �������� �ؽ�Ʈ�� ��ȯ�ϴ� �Լ�
' �Ű����� : ��ȭ������ â����(Title) , �ʱ� ������(InitialFolder), ���� ���� ���(InitialView)
'------------------------------------------------------------------------------------------------
Private Function fnGetDirectory(Title As String, Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = msoFileDialogViewList) As String
Attribute fnGetDirectory.VB_Description = "���� ã�� ��ȭ���ڸ� ���� �������� ����� ǥ���� �� ������ �������� �ؽ�Ʈ�� ��ȯ�ϴ� �Լ�"
Attribute fnGetDirectory.VB_ProcData.VB_Invoke_Func = " \n17"
    Dim V As Variant
    Dim InitFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then InitFolder = InitFolder & "\"
                
                .InitialFileName = InitFolder
            End If
        End If
        
        .Show
        
         If .SelectedItems.Count > 0 Then
             fnGetDirectory = .SelectedItems(1)
         Else
             fnGetDirectory = vbNullString
         End If
    End With
End Function

'------------------------------------------------------------------------------------------------
' ��� :  ���� ã�� ��ȭ���ڸ� ���� �������� ����� ǥ���� ��
'          ������ ������ ����� ���ϰ� �������� ������ �� ���չ����� ǥ��
'------------------------------------------------------------------------------------------------

Sub sb���ϸ��ǥ��()
Attribute sb���ϸ��ǥ��.VB_Description = "������ ���� ���� ������ �ؽ�Ʈ�� ǥ��"
Attribute sb���ϸ��ǥ��.VB_ProcData.VB_Invoke_Func = " \n17"
   Dim strPath  As String, strFile  As String
   Dim rngWork As Range, wkB As Workbook
   Dim R As Long
   
'On Error GoTo Err_Rtn
'   strPath = fnGetDirectory("�۾� ���� ����", Left(ThisWorkbook.path, 3), msoFileDialogViewSmallIcons)
   strPath = fnGetDirectory("�۾� ���� ����", , msoFileDialogViewSmallIcons)
   If strPath = "" Then Exit Sub  '<���> ���߸� ���� ���
   If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
   
   strFile = Dir(strPath, vbDirectory)
   If strFile = "" Then
      MsgBox "�������� �ʴ� ��θ��Դϴ�.", , "��θ� ����"
      Exit Sub
   End If
   
   Set wkB = Workbooks.Add
   Set rngWork = wkB.Sheets(1).Range("A1")
   
   With rngWork
      .Offset(R, 0) = "�� �۾� ���� : "
      .Offset(R, 1) = strPath
      
      R = R + 2
      .Offset(R, 0) = "���ϸ�"
      .Offset(R, 1) = "Ÿ��"
      .Offset(R, 2) = "����ũ��"
      .Offset(R, 3) = "�ۼ���"
      
      Do While strFile <> ""
         If strFile <> "." And strFile <> ".." Then
            R = R + 1
             If (GetAttr(strPath & strFile) And vbDirectory) = vbDirectory Then
                .Offset(R, 1) = "Folder"
            Else
                .Offset(R, 1) = "File"
            End If
            .Offset(R, 0) = strFile
            .Offset(R, 2) = FileLen(strPath & strFile)
            .Offset(R, 3) = FileDateTime(strPath & strFile)
         End If
         
         strFile = Dir
      Loop
   End With
  
'--// ǥ�� �����Ͽ� ���� �����ϰ� ���� ó��
   With wkB.Sheets(1).ListObjects.Add(xlSrcRange, Range("$A$3").CurrentRegion, , xlYes)
      .Name = "tblFileList"
      .TableStyle = "TableStyleMedium1"
      .ListColumns(3).Range.NumberFormatLocal = "#,##0_)"
      .ListColumns(4).Range.NumberFormatLocal = "yyyy-mm-dd hh:mm"
      Range("tblFileList[#All]").Columns.AutoFit
      
      .Sort.SortFields.Clear
      .Sort.SortFields.Add Key:=Range("tblFileList[[#All],[���ϸ�]]"), _
                  SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .Sort.Header = xlYes
      .Sort.MatchCase = False
      .Sort.Orientation = xlTopToBottom
      .Sort.SortMethod = xlPinYin
      .Sort.Apply
      .Unlist        '--// ǥ�� �ٽ� '������ ��ȯ'
   End With

Err_Rtn:
   If Err.Number <> 0 Then
      MsgBox "���� ��� �ۼ��� ���� ������ �߻��߽��ϴ�." _
            & vbCr & Err.Description, vbCritical, "�����߻�"
   Else
      MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, "�Ϸ�"
   End If
End Sub
