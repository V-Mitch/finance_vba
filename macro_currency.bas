Attribute VB_Name = "Module1"
Sub Currency_Retrieve()
Attribute Currency_Retrieve.VB_ProcData.VB_Invoke_Func = "Q\n14"
'V-Mitch
    
    'Deletion of existing query with same name and addition of a sheet
        Sheets.Add After:=ActiveSheet
        ActiveWorkbook.Queries("Table 0").Delete
        
    
    ActiveWorkbook.Queries.Add Name:="Table 0", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Web.Page(Web.Contents(""https://www.bloomberg.com/markets/currencies""))," & Chr(13) & "" & Chr(10) & "    Data0 = Source{0}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data0,{{""Currency"", type text}, {""Value"", type number}, {""Change"", type number}, {""Net Change"", Percentage.Type}, {""Time (EDT)"", type date}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ' ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Table 0"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table 0]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        ' .ListObject.DisplayName = "Table_1"
        .Refresh BackgroundQuery:=False
        
    Range("$D$1:$D$7").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("A1:A7,D1:D7").Select
    'Range("Table_ExternalData_14[[#Headers],[Net Change]]").Activate
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("$A$1:$A$7,$D$1:$D$7")
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlLow
    
    ActiveChart.FullSeriesCollection(1).Select
    Selection.InvertIfNegative = True
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(1).InvertColor = RGB(237, 138, 81)
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.ChartGroups(1).GapWidth = 46
    Range("I5").Select
    
    End With
End Sub
