# excel_vba
엑셀을 활용한 업무자동화

## advanced filter
```vb
Sub main_advanced_filter(Sub main_advanced_filter()
    Sheets("filter").Range("a9:d10000").ClearFormats
    Call advancedFilter
    정렬조건 = Sheets("filter").Range("b5").Value
    Call sort(정렬조건)

    'Sheets("report").Range("a1").Value = Sheets("filter").Range("a8").CurrentRegion

End Sub
Sub advancedFilter()
    Sheets("list").Range("A1").CurrentRegion.advancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("filter").Range("A1").CurrentRegion, CopyToRange:=Range( _
        "filter!Extract"), Unique:=False
End Sub
Sub sort(p정렬조건)
    ActiveWorkbook.Worksheets("filter").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("filter").sort.SortFields.Add Key:=Range("B9:B47") _
        , SortOn:=xlSortOnValues, Order:=p정렬조건, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("filter").sort
        .SetRange Range("A8").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

    Sheets("filter").Range("a9:d10000").ClearFormats
    Call advancedFilter
    정렬조건 = Sheets("filter").Range("b5").Value
    Call sort(정렬조건)

    'Sheets("report").Range("a1").Value = Sheets("filter").Range("a8").CurrentRegion

End Sub
Sub advancedFilter()
    Sheets("list").Range("A1").CurrentRegion.advancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Sheets("filter").Range("A1").CurrentRegion, CopyToRange:=Range( _
        "filter!Extract"), Unique:=False
End Sub
Sub sort(p정렬조건)
    ActiveWorkbook.Worksheets("filter").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("filter").sort.SortFields.Add Key:=Range("B9:B47") _
        , SortOn:=xlSortOnValues, Order:=p정렬조건, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("filter").sort
        .SetRange Range("A8").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

```

### A1셀에 '안녕하세요' 출력하는 서브루틴(subroutine)만들기
    * 서브루틴이란? ALT+F11
    * VBE(Visual Basic Editor)
    * 서브루틴 만들기
    * 서브루틴 호출하기 Call <서브루틴 이름>

### 파라메터(parameter)를 이용해 A2셀에 메세지를 받아서 출력하기
    * 파라메터란?
    * 파라메터를 사용하는 이유
    * 파라메터 여러개인 서브루틴 만들기

### 변수와 상수
    * 변수는 변하는 수(값) ""를 쓰지 않는다
    * 상수는 항상 같은 수(값) ""를 쓴다.

### 글자 연결하기
    * ex) 월 & "월"
    * & 이 연결 하는 기호이다.
    * & 의 앞뒤로는 반드시 띄어쓰기를 넣어준다.

### 주석처리 하기
    * ex) '1~20까지 출력 하는 코드 입니다.
    * 주석이란? 소스코드에서 실행이 안되는 부분(설명을 쓸때 사용한다)
    * '를 앞에 붙이면 된다

### 지우기
    * 전체 : cells.clear
    * 부분: range("a1:b6").clear

### 숫자를 영문으로 바꾸기
    * chr(65)

### function(펑션, 함수)이란?
    * return 값이 있는 subroutine

### getCellValue() function만들기
    * Excute4Macro()


