Attribute VB_Name = "mod같은색계산"
'------------------------------------------------------------------------------------------
'   【 길벗 매크로와 VBA / 도서출판 길벗 / 이동숙(bofb@naver.com) 】
'   기능 :  계산 범위와 글꼴 색이 지정된 셀을 지정해 주면
'            계산 범위에서 특정 색의 셀 값의 합계만 계산
'------------------------------------------------------------------------------------------

Function fn같은색계산(계산범위 As Range, 대상색셀 As Range, Optional 함수이름 As String = "SUM")
Attribute fn같은색계산.VB_Description = "계산 범위와 글꼴 색이 지정된 셀, 계산할 함수이름\n(예; SUM, AVERAGE, COUNT, COUNTA, MAX, MIN 등)을\n지정해 주면 범위 안에서 지정한 글꼴 색과 같은 셀만 찾아\n해당 함수로 계산한 결과를 반환합니다."
Attribute fn같은색계산.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim K As Range, rngCal As Range
   Dim Result As Double
   
   For Each K In 계산범위
      If K.Font.Color = 대상색셀.Font.Color Then
         If rngCal Is Nothing Then
            Set rngCal = K
         Else
            Set rngCal = Union(rngCal, K)
         End If
      End If
   Next
   If Not rngCal Is Nothing Then
      Result = Application.Evaluate(함수이름 & "(" & rngCal.Address & ")")
   End If
   fn같은색계산 = Result
End Function

