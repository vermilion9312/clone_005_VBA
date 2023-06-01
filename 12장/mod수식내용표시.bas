Attribute VB_Name = "mod수식내용표시"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  지정한 셀의 수식 내용을 텍스트로 표시
'------------------------------------------------------------------------------------------
Function fn수식보기(셀 As Range)
Attribute fn수식보기.VB_Description = "지정한 셀의 수식 내용을 텍스트로 표시"
Attribute fn수식보기.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim strTemp As String

    If 셀.HasArray Then
        strTemp = "{" & 셀.Formula & "}"
    ElseIf 셀.HasFormula Then
        strTemp = 셀.Formula
    Else
        strTemp = ""
    End If
    fn수식보기 = strTemp
End Function


