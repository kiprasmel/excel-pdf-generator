Option Explicit ' Declare variables as Dim before use

Sub GeneratePDFsFromCSV()
    On Error GoTo ErrorHandler ' Enable error handling

    Dim wsData As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim i As Long
	Dim j As Long
    Dim pdfPath As String
    Dim pdfName As String
    Dim pdfFullPath As String

    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Set references to worksheets
    Set wsData = ThisWorkbook.Sheets("Duomenys") ' Change "Duomenys" to your CSV data sheet name
    Set wsTemplate = ThisWorkbook.Sheets("Sablonas") ' Change "Sablonas" to your template sheet name

    ' Find the last row with data in the CSV data sheet
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' Set the PDF path to the workbook's directory with "generated-" prefix, current timestamp, and Excel filename.
    pdfPath = ThisWorkbook.Path & "\" & "generated-" & Format(Now, "yyyy-MM-dd__HH-mm-ss") & "--" & ThisWorkbook.Name & "\"

    ' Check if the folder exists, and create it if not
    If Len(Dir(pdfPath, vbDirectory)) = 0 Then
        MkDir pdfPath
    End If

    ' Loop through rows in the CSV data sheet (starting from row 2, assuming row 1 contains headers)
    For i = 2 To lastRow
        ' Copy data from CSV data sheet to the template sheet
        For j = 1 To wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
            wsTemplate.Range(wsData.Cells(1, j).Value).Value = wsData.Cells(i, j).Value
        Next j

        ' Set the PDF name. Assuming the ID is in the first column
        pdfName = (i - 1) & " " & wsData.Cells(i, 1).Value & ".pdf"

        ' Save the template sheet as a PDF in the timestamped subfolder
        pdfFullPath = pdfPath & pdfName
        wsTemplate.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFullPath
    Next i

    ' Enable screen updating back
    Application.ScreenUpdating = True

    ' Display a message when the process is complete
    MsgBox "PDFs generated successfully!", vbInformation

    Exit Sub ' Exit the subroutine if everything executed successfully

ErrorHandler:
    ' Handle errors here (e.g., display a message or log the error)
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True ' Ensure screen updating is enabled even in case of an error
End Sub

