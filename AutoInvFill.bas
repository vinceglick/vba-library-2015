Sub AutoInvFill()
'
' Macro1 Macro


'Add a "Temp" Sheet after the last Sheet in the Invoices Query File
Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Temp"

'Set the working directory for where the Export Data File (ED) resides
'Remember to enter a "\" as the last character in the directory string!!
OpenPath = "C:\Users\vglick\Documents\Macro\"

'Call the data exported from internet and assign it to a variable
ED = OpenPath & "2015-8-26" & ".xlsm"
    'Activate the ED Workbook
    Workbooks.Open(ED).Activate
        
        'Select the first row in the ED workbook and remove the apostrophe from specified header
        Rows(1).Select
        For Each currentcell In Selection
            If currentcell.Value = "'NTPEP Number" Then
                currentcell.Value = "NTPEP Number"
            End If
        Next
    
        'Find the specified header value then select & copy the values underneath
        Dim rngNTPEPnum As Range
        Set rngNTPEPnum = Range("A1:Z1").Find("NTPEP Number")
        If rngNTPEPnum Is Nothing Then
            MsgBox "NTPEP Number column was not found."
            Exit Sub
        End If
        Range(rngNTPEPnum(2), rngNTPEPnum.End(xlDown)).Select
        
        Application.Selection.Copy
        
    'Activate the workbook with which you wish to import your selected data into
    Windows("Invoices Query (2011-2015)").Activate
        
        'Add a "Temp" Sheet after the last Sheet in the Invoices Query File
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = "Temp"
        ActiveSheet.Paste
        
        'Filter out blanks in the Product Category Column and Sort the Invoices Sheet at levels "Product Catefory" and "NTPEP."
        
        Sheets("Invoices 15").Select
            ActiveSheet.Range("$A$1:$BR$873").AutoFilter Field:=2, Criteria1:="<>"
            Range("B1").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            ActiveWorkbook.Worksheets("Invoices 15").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Invoices 15").Sort.SortFields.Add Key:=Range( _
            "B2:B747"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            ActiveWorkbook.Worksheets("Invoices 15").Sort.SortFields.Add Key:=Range( _
            "D2:D747"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("Invoices 15").Sort
                .SetRange Range("B1:P747")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

        
        'Paste the NTPEP Numbers into first empty row, cell of column "NTPEP."
        
        Sheets("Temp").Select
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.Selection.Copy
        
        
        Sheets("Invoices 15").Select
        Columns("B").Find("", Cells(Rows.Count, "B")).Select
        ActiveCell.Offset(0, 2).Select
        ActiveSheet.Paste
        
        'Vlookup remaining values
        ActiveCell.Offset(0, 1).Select
        
        
'Delete Temp Sheet

Sheets("Temp").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Invoices 15").Select


'Close ED Workbook


End Sub






