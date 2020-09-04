Sub Calculo_lote()
'
' Calculo_lote Macro
' Calculo de columna Lote en hoja de Datos
'
' Acceso directo: CTRL+l
'
    Range("E6").Select
    ActiveWorkbook.RefreshAll
    Sheets("Produccion").Select
    ActiveWorkbook.Worksheets("Produccion").ListObjects("produccion").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Produccion").ListObjects("produccion").Sort. _
        SortFields.Add Key:=Range("produccion[[#All],[Fecha Produccion]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Produccion").ListObjects("produccion").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Datos").Select
    Range("S3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=IF([@Refilada],IF(R[-1]C[-5],1,R[-1]C+1),0)"
    Range("S3").Select
    Selection.AutoFill Destination:=Range("S3:S189")
    Range("S3:S189").Select
    ActiveWindow.SmallScroll Down:=172
End Sub
