Sub main_file()
    result = plus(10, 20)
    Range("a1").Value = result
    Range("a2").Value = minus(10, 20)
    Range("a3").Value = 10 ^ 2
    Range("a5").Value = 계산_평to미터제곱(8)
End Sub
'사용자가 만든 함수
Function plus(p_a, p_b)
    plus = p_a + p_b
End Function

Function minus(p_a, p_b)
    minus = p_a - p_b
End Function

Function power(p_a)

End Function

Function 계산_평to미터제곱(p_평)
    계산_평to미터제곱 = p_평 * 3.3058
End Function
