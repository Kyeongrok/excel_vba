Sub main()
    Call 출력_안녕하세요
    Call 출력_메세지("안녕히가세요")
    Call 출력_메세지("반갑습니다.")
    Call 출력_메세지_범위("또 뵙겠습니다.", "a3")
    Call 출력_메세지_범위("또 뵙겠습.", "a4")
    Call 출력_메세지_범위("안녕하세요 김경록님", "a5")
    Call 출력_메세지_범위("안녕히가세요 김경록님", "a6")
    Call 출력_메세지_범위_이름("안녕히가세요", "a7", "김경록")
    Call 출력_메세지_범위_이름("안녕히가세요", "a8", "김영환")

End Sub

'서브루틴 sub
'excel vba 명령 실행 단위
'Parameter가 0개인 서브루틴
Sub 출력_안녕하세요()
    Range("a1").Value = "안녕하세요"
End Sub

'parameter 파라메터 매개변수
'parameter가 1개인 서브루틴
Sub 출력_메세지(p_메세지)
    Range("a2").Value = p_메세지
End Sub

'변수 변하는 수(값) ""를 안쓴다
'상수 항상 같은 수(값) ""를 쓴다
'parameter가 2개인 서브루틴
Sub 출력_메세지_범위(p_메세지, p_범위)
    Range(p_범위).Value = p_메세지
End Sub

'parameter가 3개인 서브루틴
Sub 출력_메세지_범위_이름(p_메세지, p_범위, p_이름)
    Range(p_범위).Value = p_메세지 & " " & p_이름
End Sub