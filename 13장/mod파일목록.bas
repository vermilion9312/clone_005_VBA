Attribute VB_Name = "mod파일목록"
Option Explicit
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  파일 찾기 대화상자를 열어 폴더명의 목록을 표시한 후
'          선택한 폴더에 저장된 파일과 폴더들의 정보를 새 통합문서에 표시
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
' 기능 :  파일 찾기 대화상자를 열어 폴더명의 목록을 표시한 후
'          선택한 폴더명을 텍스트로 반환하는 함수
' 매개변수 : 대화상자의 창제목(Title) , 초기 폴더명(InitialFolder), 파일 보기 방식(InitialView)
'------------------------------------------------------------------------------------------------
Private Function fnGetDirectory(Title As String, Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = msoFileDialogViewList) As String
Attribute fnGetDirectory.VB_Description = "파일 찾기 대화상자를 열어 폴더명의 목록을 표시한 후 선택한 폴더명을 텍스트로 반환하는 함수"
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
' 기능 :  파일 찾기 대화상자를 열어 폴더명의 목록을 표시한 후
'          선택한 폴더에 저장된 파일과 폴더들의 정보를 새 통합문서에 표시
'------------------------------------------------------------------------------------------------

Sub sb파일목록표시()
Attribute sb파일목록표시.VB_Description = "지정한 셀의 수식 내용을 텍스트로 표시"
Attribute sb파일목록표시.VB_ProcData.VB_Invoke_Func = " \n17"
   Dim strPath  As String, strFile  As String
   Dim rngWork As Range, wkB As Workbook
   Dim R As Long
   
'On Error GoTo Err_Rtn
'   strPath = fnGetDirectory("작업 폴더 선택", Left(ThisWorkbook.path, 3), msoFileDialogViewSmallIcons)
   strPath = fnGetDirectory("작업 폴더 선택", , msoFileDialogViewSmallIcons)
   If strPath = "" Then Exit Sub  '<취소> 단추를 누른 경우
   If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
   
   strFile = Dir(strPath, vbDirectory)
   If strFile = "" Then
      MsgBox "존재하지 않는 경로명입니다.", , "경로명 오류"
      Exit Sub
   End If
   
   Set wkB = Workbooks.Add
   Set rngWork = wkB.Sheets(1).Range("A1")
   
   With rngWork
      .Offset(R, 0) = "▣ 작업 폴더 : "
      .Offset(R, 1) = strPath
      
      R = R + 2
      .Offset(R, 0) = "파일명"
      .Offset(R, 1) = "타입"
      .Offset(R, 2) = "파일크기"
      .Offset(R, 3) = "작성일"
      
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
  
'--// 표로 지정하여 서식 지정하고 정렬 처리
   With wkB.Sheets(1).ListObjects.Add(xlSrcRange, Range("$A$3").CurrentRegion, , xlYes)
      .Name = "tblFileList"
      .TableStyle = "TableStyleMedium1"
      .ListColumns(3).Range.NumberFormatLocal = "#,##0_)"
      .ListColumns(4).Range.NumberFormatLocal = "yyyy-mm-dd hh:mm"
      Range("tblFileList[#All]").Columns.AutoFit
      
      .Sort.SortFields.Clear
      .Sort.SortFields.Add Key:=Range("tblFileList[[#All],[파일명]]"), _
                  SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .Sort.Header = xlYes
      .Sort.MatchCase = False
      .Sort.Orientation = xlTopToBottom
      .Sort.SortMethod = xlPinYin
      .Sort.Apply
      .Unlist        '--// 표를 다시 '범위로 변환'
   End With

Err_Rtn:
   If Err.Number <> 0 Then
      MsgBox "파일 목록 작성중 다음 오류가 발생했습니다." _
            & vbCr & Err.Description, vbCritical, "오류발생"
   Else
      MsgBox "작업이 완료되었습니다.", vbInformation, "완료"
   End If
End Sub
