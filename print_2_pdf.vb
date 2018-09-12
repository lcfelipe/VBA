Sub PrintToPDF_MultiSheetToOne_Early()
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Print to PDF file using PDFCreator
'   (Download from http://sourceforge.net/projects/pdfcreator/)
'   Designed for early bind, set reference to PDFCreator
    Dim pdfjob As PDFCreator.clsPDFCreator
    Dim sPDFName As String
    Dim sPDFPath As String
    Dim lSheet As Long
    Dim lTtlSheets As Long
    '/// Change the output file name here! ///
    sPDFName = "Consolidated.pdf"
    sPDFPath = ActiveWorkbook.path & Application.PathSeparator
    Set pdfjob = New PDFCreator.clsPDFCreator
    'Make sure the PDF printer can start
    If pdfjob.cStart("/NoProcessingAtStartup") = False Then
        MsgBox "Can't initialize PDFCreator.", vbCritical + _
                vbOKOnly, "Error!"
        Exit Sub
    End If
    'Set all defaults
    With pdfjob
        .cOption("UseAutosave") = 1
        .cOption("UseAutosaveDirectory") = 1
        .cOption("AutosaveDirectory") = sPDFPath
        .cOption("AutosaveFilename") = sPDFName
        .cOption("AutosaveFormat") = 0    ' 0 = PDF
        .cClearCache
    End With
    'Print the document to PDF
    lTtlSheets = Application.Sheets.Count
    For lSheet = 1 To Application.Sheets.Count
        On Error Resume Next 'To deal with chart sheets
        If Not IsEmpty(Application.Sheets(lSheet).UsedRange) Then
            Application.Sheets(lSheet).PrintOut copies:=1, ActivePrinter:="PDFCreator"
        Else
            lTtlSheets = lTtlSheets - 1
        End If
        On Error GoTo 0
    Next lSheet
    'Wait until all print jobs have entered the print queue
    Do Until pdfjob.cCountOfPrintjobs = lTtlSheets
        DoEvents
    Loop
    'Combine all PDFs into a single file and stop the printer
    With pdfjob
        .cCombineAll
        .cPrinterStop = False
    End With
    'Wait until PDF creator is finished then release the objects
    Do Until pdfjob.cCountOfPrintjobs = 0
        DoEvents
    Loop
    pdfjob.cClose
    Set pdfjob = Nothing
End Sub