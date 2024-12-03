'Globals
Global newWsName As String, totalPages As Integer, globalLog As String, cp As Integer, dict As Object



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
    
        
    
    'Combine pages

    'Loop through Sheets
    Do While cp < totalPages
    
        'increment
        cp = cp + 1
        
        
        Sheets("Page " & cp).Activate
        'Range("U3").Select
        Debug.Print Range("U3")
        
        
        'DebugMsg ("Page " & cp & " is good.")
        
        
        
        'Find Invoice Numbers
        
        
        
        
        'Return array/dictionary of invoice numbers
        
        
        
        
        'Output array/dictionary to Invoices
        
        
        
        
        
        
        
        'Columns("AR:BB").Select
        


    Loop
    
    
    'Create new compiled page of invoices
    newWsName = "Invoices"
    CompileInvoices (newWsName)




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

Sub CompileInvoices(wsName)


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
