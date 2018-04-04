Sub main_chart()
    Sheets("Sheet1").ChartObjects.Delete
    Call drawChart
End Sub
Sub drawChart()
    Sheets("Sheet1").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$A$27,Sheet1!$C$1:$C$27" _
        )
End Sub