Attribute VB_Name = "modPictureSaveAPI"
Option Explicit
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  특정 영역을 그림으로 저장하는 기능
'   API는 오피스 설치 버전(2010 이상)에 따라 아래와 같이 다르게 사용해야 함
'------------------------------------------------------------------------------------------
#If Win64 And VBA7 Then
      Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
      Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
      Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
      Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
      Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
      Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As LongPtr) As Long

#Else
      Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
      Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
      Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
      Private Declare Function CloseClipboard Lib "user32" () As Long
      Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
      Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
#End If

Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4

'------------------------------------------------------------------------------------------
' 클립보드의 내용을 그림으로 저장하고 클립보드의 내용을 지움(닫음)
'------------------------------------------------------------------------------------------
Private Function sbSavePicture(strFileName As String, Optional lXlPicType As Long = xlPicture) As Boolean
   Dim h As Long, hPicAvail As Long, hPtr As Long, hPal As Long, lPicType As Long, hCopy As Long
   Dim hClose As Long
   
   sbSavePicture = False
   
   lPicType = IIf(lXlPicType = xlBitmap, CF_BITMAP, CF_ENHMETAFILE)
   
   '클립보드에 그림 형식이 저장되었는지 확인
   hPicAvail = IsClipboardFormatAvailable(lPicType)
   
   If hPicAvail <> 0 Then
       h = OpenClipboard(0&) '클립보드를 열고
       If h > 0 Then
           hPtr = GetClipboardData(lPicType) '클립보드의 데이터를 얻어
           hCopy = CopyEnhMetaFile(hPtr, strFileName) '메타파일로 저장합니다
           If hCopy Then
               DeleteEnhMetaFile hCopy '메타파일에 관련한 시스템 리스스를 해제
               sbSavePicture = True
           End If
           h = CloseClipboard '클립보드를 닫습니다
       End If
   End If
End Function

'------------------------------------------------------------------------------------------
' 기능 : 셀 영역을 그림으로 저장
' 매개변수 : rngDB - 그림으로 저장할 셀 영역 , strFileName - 저장할 경로명
'------------------------------------------------------------------------------------------
Public Sub sbSavePic(rngDb As Range, strFileName As String)
   If rngDb.Areas.Count = 1 Then
       rngDb.CopyPicture xlScreen, xlPicture    '--// 현재 화면상태를 그림으로 복사(클립보드로 복사됨)
       sbSavePicture strFileName, xlPicture     '--//  클립보드의 내용을 그림으로 지정한 파일명으로 저장
   End If
End Sub


