Attribute VB_Name = "Module2"
'Globals
Global newWsName As String, totalPages As Integer, globalLog As String, cp As Integer, dict As Object, aSheet As String



Public Sub AgingRpt90()
'
' AgingRpt90 Macro

    'Initialize Error Log
    globalLog = vbNewLine & vbNewLine & "Error Log for Macro " & Format(Now(), "MMM-DD-YYYY") & " :"
    
    
    'Initialize dictionary
    'Late Binding
    Set dict = CreateObject("Scripting.Dictionary")
    
    
    'Initializing count
    cp = 0
    
    

    ' Get total number of workbook pages
    totalPages = ThisWorkbook.Sheets.Count
    dict.Add Key:="TotalPages", Item:=totalPages

    'MsgBox dict("TotalPages")
    
        
    
    'Get page info

    'Loop through Sheets
    Do While cp < totalPages
    
    
        'increment
        cp = cp + 1
        
        
        aSheet = "Page " & cp
        
        
        Sheets(aSheet).Activate
        'Range("U3").Select
        'DebugMsg (aSheet)
        
        'Remove blank columns and rows
        DeleteEmptyColumnsForTable (aSheet)
        DeleteEmptyRowsForTable (aSheet)
        

        
        
        'DebugMsg ("Page " & cp & " is good.")
        
        
        'Find Invoice Numbers
        'TODO
        GetInvoiceData (cp)
        
        
        
        'Return array/dictionary of invoice numbers
        'TODO
        
        
        
        
        
        
        'Columns("AR:BB").Select
        


    Loop
    
    
    'Create new compiled page of invoices
    newWsName = "Invoices"
    CompileInvoices (newWsName)
    
    
    'Output array/dictionary to Invoices
    'TODO
    PutInvoiceData (newWsName)




    'Columns("AR:BB").Select
    'Range("BB1").Activate
    'Selection.Delete Shift:=xlToLeft
    
    
    'Sheets("Page 1").Select
    'Columns("BL:BN").Select
    'Selection.Delete Shift:=xlToLeft
    'Columns("BJ:BJ").ColumnWidth = 9
    'Columns("BJ:BJ").ColumnWidth = 42.29
    'Columns("BJ:BL").Select
    'Selection.Delete Shift:=xlToLeft
    'Columns("BJ:BP").Select



    'Output log
    OutputLogFile
    
    
    
End Sub



Sub CompileInvoices(wsName As String)


    'If workbook exist skip
    If cp > dict("TotalPages") Then
    
        'Error
        DebugMsg (wsName & " is a page name that already exists.")
        
    Else
    
        'Create new Workbook
        Sheets.Add.Name = wsName
        
        Sheets(wsName).Activate
        
        'Add header columns
        WriteInvoicesHeaders (wsName)
        
        
    End If
    

End Sub



Sub WriteInvoicesHeaders(pageWrites As String)


    Sheets(pageWrites).Range("A1").Value = "Project ID/Cost Center"

    Sheets(pageWrites).Range("B1").Value = "Invoice #"

    Sheets(pageWrites).Range("C1").Value = "Ref. No."

    Sheets(pageWrites).Range("D1").Value = "Invoice Data"

    Sheets(pageWrites).Range("E1").Value = "Student"

    Sheets(pageWrites).Range("F1").Value = "Course #"

    Sheets(pageWrites).Range("G1").Value = "Current"

    Sheets(pageWrites).Range("H1").Value = "Over 90 days past due"
    
    'Headers logged
    DebugMsg ("Invoices headers created")

    
End Sub





Sub GetInvoiceData(cp As Integer)

    'Declare last row and last column
    Dim lastCol As Integer, lastRow As Integer
    Dim r As Range

    DebugMsg ("Get invoice page: " & cp)
    
    
    'finds the last column
    
    lastCol = ActiveSheet.UsedRange.Columns.Count
    
    DebugMsg ("Invoice page: " & cp & " Last Column: " & lastCol)
    
    
    
    'finds the last row

    lastRow = ActiveSheet.UsedRange.Rows.Count
    
    DebugMsg ("Invoice page: " & cp & " Last Row: " & lastRow)
    

    
    'sets the range

    'Set r = Range("A1", Cells(lastRow, lastCol))
    
    
    'Remove header
    'DeleteHeaderForTable (cp, lastCol, lastRow)
    
    

End Sub




Sub DeleteEmptyRowsForTable(aSheet)

    Dim lRow As Integer
    
    Sheets(aSheet).Activate
    
    'finds the last row
    lRow = ActiveSheet.UsedRange.Rows.Count
    
    
    If Range("A1:" & lRow & "1").SpecialCells(xlCellTypeBlanks).EntireRow.Count > 0 Then
    
        'Delete empty rows
        DebugMsg ("Deleting :" & Range("A1:" & lRow & "1").SpecialCells(xlCellTypeBlanks).EntireRow.Count & " rows")
        Range("A1:" & lRow & "1").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Else
        DebugMsg ("No rows to delete on Page: " & aSheet)
        
        
    End If

End Sub




Sub DeleteEmptyColumnsForTable(aSheet)

    Dim lCol As Integer, msg As String
        
    'Sheets(activeSheet).Activate
    
    'finds the last row
    lCol = ActiveSheet.UsedRange.Columns.Count
    


    
    If Range("A1:A" & lCol).SpecialCells(xlCellTypeBlanks).EntireColumn.Count > 0 Then
    
        'Delete empty columns
        DebugMsg ("Deleting :" & Range("A1:A" & lCol).SpecialCells(xlCellTypeBlanks).EntireColumn.Count & " columns")

        Range("A1:A" & lCol).SpecialCells(xlCellTypeBlanks).EntireColumn.Delete
    
    Else
        DebugMsg ("No columns to delete on Page: " & aSheet)
        
        
    End If

End Sub




Sub DeleteHeaderForTable(cp, lCol As Integer, lRow As Integer)

    'Delete header
    'Find Invoice No.
    
    
    'Delete range cell of Invoice No. and above
    'Range("A1:" & lcol & lrow ).EntireRow.Delete
    

End Sub


Sub PutInvoiceData(wsPage As String)
    
    'Headers logged
    DebugMsg ("Output " & wsPage & " for " & dict("TotalPages") & " pages.")


End Sub



Sub DebugMsg(infoMsg As String)

    'Print message to Immediate console
    Debug.Print "FYI:  " & infoMsg
    
    'Add to log
    Lumberjack (infoMsg)
    
    
End Sub



Sub Lumberjack(logging As String)
    
    globalLog = globalLog & vbNewLine & vbNewLine & logging

End Sub



Sub OnErrorStatement(cp)


ErrorHandler:
    
    'Add to log
    Lumberjack ("Page " & cp & " has a problem.")

End Sub



Sub OutputLogFile()

    Dim path As String, fileNumber As Integer
        
    
    'Set save path
    path = "C:\Users\musslema\OneDrive - University of Texas at Arlington\Desktop\MacroLog.txt"
    
    'Saves File Number/pointer
    fileNumber = FreeFile
    
    Open path For Append As fileNumber

        Print #fileNumber, globalLog
    
        Close fileNumber
    
        Shell "notepad.exe " & path, vbNormalFocus

End Sub

