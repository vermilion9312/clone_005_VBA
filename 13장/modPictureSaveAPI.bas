Attribute VB_Name = "modPictureSaveAPI"
Option Explicit
'------------------------------------------------------------------------------------------
'   �� ��� ��ũ�ο� VBA / �������� ��� / �̵���(bofb@naver.com) ��
'   ��� :  Ư�� ������ �׸����� �����ϴ� ���
'   API�� ���ǽ� ��ġ ����(2010 �̻�)�� ���� �Ʒ��� ���� �ٸ��� ����ؾ� ��
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
' Ŭ�������� ������ �׸����� �����ϰ� Ŭ�������� ������ ����(����)
'------------------------------------------------------------------------------------------
Private Function sbSavePicture(strFileName As String, Optional lXlPicType As Long = xlPicture) As Boolean
   Dim h As Long, hPicAvail As Long, hPtr As Long, hPal As Long, lPicType As Long, hCopy As Long
   Dim hClose As Long
   
   sbSavePicture = False
   
   lPicType = IIf(lXlPicType = xlBitmap, CF_BITMAP, CF_ENHMETAFILE)
   
   'Ŭ�����忡 �׸� ������ ����Ǿ����� Ȯ��
   hPicAvail = IsClipboardFormatAvailable(lPicType)
   
   If hPicAvail <> 0 Then
       h = OpenClipboard(0&) 'Ŭ�����带 ����
       If h > 0 Then
           hPtr = GetClipboardData(lPicType) 'Ŭ�������� �����͸� ���
           hCopy = CopyEnhMetaFile(hPtr, strFileName) '��Ÿ���Ϸ� �����մϴ�
           If hCopy Then
               DeleteEnhMetaFile hCopy '��Ÿ���Ͽ� ������ �ý��� �������� ����
               sbSavePicture = True
           End If
           h = CloseClipboard 'Ŭ�����带 �ݽ��ϴ�
       End If
   End If
End Function

'------------------------------------------------------------------------------------------
' ��� : �� ������ �׸����� ����
' �Ű����� : rngDB - �׸����� ������ �� ���� , strFileName - ������ ��θ�
'------------------------------------------------------------------------------------------
Public Sub sbSavePic(rngDb As Range, strFileName As String)
   If rngDb.Areas.Count = 1 Then
       rngDb.CopyPicture xlScreen, xlPicture    '--// ���� ȭ����¸� �׸����� ����(Ŭ������� �����)
       sbSavePicture strFileName, xlPicture     '--//  Ŭ�������� ������ �׸����� ������ ���ϸ����� ����
   End If
End Sub


