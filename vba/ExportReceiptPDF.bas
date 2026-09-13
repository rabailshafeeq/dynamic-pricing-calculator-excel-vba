Attribute VB_Name = "Module2"
' ============================================================
' ExportReceiptPDF
' Exports the Receipt sheet as a PDF, named using the pet's name
' and today's date, into a "Receipts" folder next to the workbook.
' Independent of SaveQuoteToLog - can be used whether or not a
' booking has been logged yet.
' ============================================================

Sub ExportReceiptPDF()

    Dim wsQuote As Worksheet, wsReceipt As Worksheet
    Dim petName As String, safeName As String, fileName As String, folderPath As String
    Dim i As Integer

    Set wsQuote = ThisWorkbook.Sheets("Pricing Calculator")
    Set wsReceipt = ThisWorkbook.Sheets("Receipt")

    petName = wsQuote.Range("F4").Value
    If petName = "" Then
        MsgBox "Please enter a Pet Name before exporting a receipt.", vbExclamation
        Exit Sub
    End If

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save this workbook first (so there's a folder to save the PDF into), then try again.", vbExclamation
        Exit Sub
    End If

    safeName = petName
    Dim badChars As String
    badChars = "\/:*?""<>|"
    For i = 1 To Len(badChars)
        safeName = Replace(safeName, Mid(badChars, i, 1), "")
    Next i

    fileName = safeName & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    folderPath = ThisWorkbook.Path & "\Receipts"

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    wsReceipt.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=folderPath & "\" & fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Receipt saved to: Receipts\" & fileName, vbInformation

End Sub
