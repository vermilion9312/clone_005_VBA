Attribute VB_Name = "mod하위폴더포함파일목록"
'-----------------------------------------------------------------
'  기  능 : 하위폴더 포함하여 파일 목록 만들기
'  작성자 :  길벗 엑셀 매크로와 VBA 바이블
'-----------------------------------------------------------------

'--// 사용자 정의 폼에 시트 목로 표시한 후 시트 이동

'--// 하위폴더를 포함하여 파일 리스트
Dim rngS As Range
Dim Cnt As Long
Dim subFolder As String

Sub sbFile_List()
  Dim i As Long
  Dim fdNm
  Dim iOk As Integer
  
  iOk = MsgBox("폴더의 파일 목록을 검색합니다." _
            & vbCrLf & "하위 폴더를 포함하여 검색 할까요?" _
            , vbQuestion + vbYesNo, "파일 목록")
  ' 구분, 폴더명, 파일명, 크기, 생성일자
  Set rngS = ThisWorkbook.Sheets(1).Range("A1")
  With rngS
    .CurrentRegion.Clear
    .Offset(0, 0) = "상위폴더"
    .Offset(0, 1) = "이름"
    .Offset(0, 2) = "구분"
    .Offset(0, 3) = "크기(Byte)"
    .Offset(0, 4) = "작성일"
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

'--// 정렬
  rngS.CurrentRegion.Columns.AutoFit
  With rngS.Parent.Sort
    .SortFields.Clear  '정렬 기준을 초기화
    .SortFields.Add Key:=rngS, Order:=xlAscending
    .SortFields.Add Key:=rngS.Offset(0, 1), Order:=xlAscending
    .SetRange rngS.CurrentRegion
    .Header = xlYes
    .Apply
  End With
    
'--// 조건부 서식
  Cells.FormatConditions.Delete
  Dim fc As FormatCondition
  
  With Range("A1").CurrentRegion
    Set fc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$C1=""Folder""")
    
    With fc.Interior
      .ThemeColor = xlThemeColorAccent6
      .TintAndShade = 0.5
    End With

  End With
    
'--// 필터 적용
  Range("A1").CurrentRegion.AutoFilter

'--// 틀고정
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

