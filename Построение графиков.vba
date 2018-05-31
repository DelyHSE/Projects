Public v_fc As Integer, v_lc As Integer, v_fr As Integer, v_lr As Integer, v_fc1 As Integer, v_lc1 As Integer, v_fr1 As Integer, v_lr1 As Integer
Const sheet1 = "Проекты", sheet2 = "Сотрудники и проекты"

Sub CallBack(ParamArray varname()) 'функция актуализации данных в датапровайдере

Dim r_dp As Range
Dim r_dp1 As Range
    
With Worksheets(sheet1) 'считывание значения ячейки с нужным инфопровайдером
    If varname(0) = "DP_1" Then
        Set r_dp = varname(1)
        Call RangeAdress(r_dp) 'вызов функции для присвоения значений глобальным переменным для сбора данных для первого графика
        Call CreateGraph 'вызов функции построения первого графика
    End If
End With

With Worksheets(sheet2) 'считывание значения ячейки с нужным инфопровайдером
    If varname(0) = "DP_2" Then
        Set r_dp1 = varname(1)
        Call RangeAdress1(r_dp1) 'вызов функции для присвоения значений глобальным переменным для сбора данных для второго графика
        Call CreateGraph1 'вызов функции построения второго графика
    End If
End With

End Sub

Sub CreateGraph()

Dim oChart As ChartObject 'инициализация переменных
Dim range1 As Range

With Worksheets(sheet1).ChartObjects 'удаление имеющихся на листе графиков
     If .Count > 0 Then .Delete
End With

With ThisWorkbook.Sheets(sheet1)
Set range1 = .Range(.Cells(v_fr + 2, 1), .Cells(v_lr - 1, v_lc)) 'задание диапазона данных для графика (fr = first row, lr - last row, fc - first column, lc - last column)
End With

a = range1.Address

Set oChart = ThisWorkbook.Sheets(sheet1).ChartObjects.Add(500, 70, 600, 600) 'определение размера и положения графика

oChart.Chart.SetSourceData (Sheets(sheet1).Range(a)) 'вставка данных в диаграмму

oChart.Chart.ChartType = xlBarClustered 'определение типа диаграммы

oChart.Activate 'активация диаграммы для работы с её свойствами
ActiveChart.Axes(xlCategory).ReversePlotOrder = True 'изменение порядка вывода
With ActiveChart
    .HasTitle = True
    .HasLegend = True
    .ChartTitle.Text = "Запланированные и фактические часы по проектам и сотрудникам" 'присвоение названия
    .HasLegend = True
    .Legend.Select
    .SeriesCollection(2).Name = "Фактические часы" 'присвоение легенды
    .SeriesCollection(1).Name = "Планируемые часы"
End With

End Sub

Sub CreateGraph1()

Dim oChart As ChartObject
Dim range2 As Range

With Worksheets(sheet2).ChartObjects
     If .Count > 0 Then .Delete
End With

With ThisWorkbook.Sheets(sheet2)
Set range2 = .Range(.Cells(v_fr1 + 2, 1), .Cells(v_lr1 - 1, v_lc1))
End With

a = range2.Address

Set oChart = ThisWorkbook.Sheets(sheet2).ChartObjects.Add(500, 70, 600, 600)

oChart.Chart.SetSourceData (Sheets(sheet2).Range(a))

oChart.Chart.ChartType = xlBarClustered

oChart.Activate
ActiveChart.Axes(xlCategory).ReversePlotOrder = True
With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "Запланированные и фактические часы по проектам и сотрудникам"
    .HasLegend = True
    .Legend.Select
    .SeriesCollection(2).Name = "Фактические часы"
    .SeriesCollection(1).Name = "Планируемые часы"
End With

End Sub

Sub RangeAdress(r_dp As Range) 'функция выбора значений из нужного диапазона датапровайдера, обозначили строки и столбцы, которые нам нужны
 v_fr = r_dp.Row
 v_lr = r_dp.Rows.Count + r_dp.Row - 1
 v_fc = r_dp.Column
 v_lc = r_dp.Columns.Count + r_dp.Column - 1
End Sub

Sub RangeAdress1(r_dp1 As Range)
 v_fr1 = r_dp1.Row
 v_lr1 = r_dp1.Rows.Count + r_dp1.Row - 1
 v_fc1 = r_dp1.Column
 v_lc1 = r_dp1.Columns.Count + r_dp1.Column - 1
End Sub




