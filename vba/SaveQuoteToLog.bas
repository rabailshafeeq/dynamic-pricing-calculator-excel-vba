Attribute VB_Name = "Module1"
' ============================================================
' SaveQuoteToLog
' Reads the current quote from the Pricing Calculator sheet and
' appends it as a new row on the Booking Log sheet, then clears
' the pet-specific inputs so the tool is ready for the next quote.
'
' Requires: sheets named exactly "Pricing Calculator" and "Booking Log",
' both protected with the password set in the pwd constant below.
' ============================================================

Sub SaveQuoteToLog()

    Dim wsQuote As Worksheet, wsLog As Worksheet
    Dim nextRow As Long
    Dim clientName As String, petName As String
    Dim addonList As String, serviceUsed As String
    Dim i As Long
    Const pwd As String = "Demo2026"
    Const lastBookingRow As Long = 24

    Set wsQuote = ThisWorkbook.Sheets("Pricing Calculator")
    Set wsLog = ThisWorkbook.Sheets("Booking Log")

    petName = wsQuote.Range("F4").Value
    If petName = "" Then
        MsgBox "Please enter a Pet Name before saving.", vbExclamation
        Exit Sub
    End If

    clientName = InputBox("Client Name:", "Save Booking")
    If clientName = "" Then Exit Sub

    wsQuote.Unprotect Password:=pwd
    wsLog.Unprotect Password:=pwd

    ' Search only the real booking area (rows 5-24) - not the whole column,
    ' which would wander into the Quick Summary labels below and mistake
    ' them for existing bookings.
    nextRow = 5
    Do While nextRow <= lastBookingRow And wsLog.Cells(nextRow, "A").Value <> ""
        nextRow = nextRow + 1
    Loop

    If nextRow > lastBookingRow Then
        MsgBox "The Booking Log is full! Please insert more rows above the " & _
               "Quick Summary section before saving another booking.", vbExclamation
        wsQuote.Protect Password:=pwd
        wsLog.Protect Password:=pwd
        Exit Sub
    End If

    If wsQuote.Range("B4").Value = "Cat" Then
        serviceUsed = wsQuote.Range("B13").Value
    Else
        serviceUsed = wsQuote.Range("B10").Value
    End If

    wsLog.Cells(nextRow, 1).Value = Date
    wsLog.Cells(nextRow, 2).Value = clientName
    wsLog.Cells(nextRow, 3).Value = petName
    wsLog.Cells(nextRow, 4).Value = wsQuote.Range("B4").Value
    wsLog.Cells(nextRow, 5).Value = wsQuote.Range("N4").Value   ' effective coat/hair type
    wsLog.Cells(nextRow, 6).Value = wsQuote.Range("B6").Value
    wsLog.Cells(nextRow, 7).Value = serviceUsed

    ' Add-ons: check for ANY non-blank marker, not one exact character -
    ' avoids silent mismatches between similar-looking Unicode symbols.
    addonList = ""
    For i = 16 To 33
        If Trim(wsQuote.Cells(i, 1).Value) <> "" Then
            If addonList <> "" Then addonList = addonList & ", "
            addonList = addonList & wsQuote.Cells(i, 2).Value
        End If
    Next i
    wsLog.Cells(nextRow, 8).Value = addonList

    If wsQuote.Range("B38").Value = "Percentage" Then
        wsLog.Cells(nextRow, 9).Value = wsQuote.Range("B39").Value & "%"
    Else
        wsLog.Cells(nextRow, 9).Value = "$" & wsQuote.Range("B39").Value & " fixed"
    End If

    wsLog.Cells(nextRow, 10).Value = wsQuote.Range("J17").Value
    wsLog.Cells(nextRow, 11).Value = "No"

    ' Note the breed automatically if one was used (N4 differs from B5 only
    ' when B5 was a breed name that got translated to a coat type)
    If wsQuote.Range("N4").Value <> wsQuote.Range("B5").Value Then
        wsLog.Cells(nextRow, 12).Value = "Breed: " & wsQuote.Range("B5").Value
    End If

    wsQuote.Range("F4").Value = ""
    wsQuote.Range("B5").Value = ""

    wsQuote.Protect Password:=pwd
    wsLog.Protect Password:=pwd

    MsgBox "Booking saved for " & petName & " (" & clientName & ")!", vbInformation

End Sub
