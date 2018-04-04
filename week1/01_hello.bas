Sub main()
    Call 출력_메세지("또 뵙겠습니다")
    Call 출력_메세지_이름("새해 복 많이받으세요", "경록")
End Sub

'파라메터가 2개인 서브루틴
Sub 출력_메세지_이름(p_메세지, p_이름)
    Range("a3").Value = p_이름 & p_메세지
End Sub

'파라메터 parameter 매개변수 파라메터가 1개인 sub
Sub 출력_메세지(p_메세지)
    Range("a2").Value = p_메세지
End Sub

'서브루틴 subroutine
Sub 출력_안녕하세요()
    Range("a1:b1").Value = "안녕하세요"

End Sub