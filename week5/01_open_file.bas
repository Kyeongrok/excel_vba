Sub main()
    Cells.Clear
    For 열 = 1 To 4
        For 행 = 1 To 37
            result = getCellValue("C:\Users\FTA\Desktop\", "예제_병원현황.xlsx", "list", 행, 열)
            Cells(행, 열).Value = result
        Next
    Next
End Sub

Function getCellValue(p_경로명, p_파일명, p_시트명, p_행, p_열)
    msg = "'" + p_경로명 + "[" + p_파일명 + "]" + p_시트명 _
    + "'!" + "R" & p_행 & "C" & p_열
    getCellValue = ExecuteExcel4Macro(msg)
End Function
