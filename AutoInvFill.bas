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
        
      'Since you are currently in a workbook that was initially active,
      'the landing sheet can be referred to as your ActiveSheet
            ActiveSheet.Paste
      'Extract the reference from the first three characters of A1 on the ActiveSheet
            Ident = Right(A1, 3)
            
            
        
        
End Sub
        
        
        'Match = Match(Ident, "Identifier", 0).Select
        
        'Range("NTPEP Number", Selection.End(xlDown)).Select
        'Selection.End(xlDown).Select
        'Range("A11").Select
        'Selection.Insert Shift:=xlDown
        'ActiveSheet.Paste
        'Application.CutCopyMode = False




