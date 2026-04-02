Attribute VB_Name = "Module2"
' OutlookEmailToMarkdown
' Copies selected Outlook emails to clipboard as structured Markdown
' for pasting into AI/LLM interfaces.
'
' Features:
'   - No FM20.DLL dependency (Win32 API clipboard access)
'   - Parses quoted reply chains into hierarchical Markdown headings
'   - Normalizes whitespace for token efficiency
'   - Strips common email signature blocks
'
' Usage: Select one or more emails in Outlook Explorer, then run macro.
' Install: Alt+F11 > Insert Module > paste this code.
'          No additional references required.

Option Explicit

' ---------------------------------------------------------------
' Win32 API declarations for clipboard access
' ---------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const GMEM_MOVEABLE = &H2
Private Const CF_UNICODETEXT = 13

' ---------------------------------------------------------------
' Copies a string to the Windows clipboard via Win32 API
' ---------------------------------------------------------------
Private Sub CopyToClipboard(strText As String)
    #If VBA7 Then
        Dim hGlobal As LongPtr
        Dim lpGlobal As LongPtr
    #Else
        Dim hGlobal As Long
        Dim lpGlobal As Long
    #End If

    Dim byteLen As Long

    ' Unicode string length in bytes + null terminator
    byteLen = (Len(strText) + 1) * 2

    hGlobal = GlobalAlloc(GMEM_MOVEABLE, byteLen)
    If hGlobal = 0 Then
        MsgBox "Failed to allocate memory for clipboard.", vbCritical
        Exit Sub
    End If

    lpGlobal = GlobalLock(hGlobal)
    If lpGlobal = 0 Then
        MsgBox "Failed to lock memory for clipboard.", vbCritical
        Exit Sub
    End If

    CopyMemory lpGlobal, StrPtr(strText), byteLen
    GlobalUnlock hGlobal

    If OpenClipboard(0) = 0 Then
        MsgBox "Failed to open clipboard.", vbCritical
        Exit Sub
    End If

    EmptyClipboard
    SetClipboardData CF_UNICODETEXT, hGlobal
    CloseClipboard
End Sub

' ---------------------------------------------------------------
' Main entry point
' ---------------------------------------------------------------
Sub CopySelectedEmailsToMarkdown()
    On Error GoTo ErrorHandler

    Dim objSelection As Selection
    Dim objMail As MailItem
    Dim strOutput As String
    Dim i As Long
    Dim mailCount As Long

    Set objSelection = Application.ActiveExplorer.Selection

    If objSelection.Count = 0 Then
        MsgBox "No emails selected.", vbExclamation
        Exit Sub
    End If

    mailCount = 0
    For i = 1 To objSelection.Count
        If TypeOf objSelection.Item(i) Is MailItem Then
            Set objMail = objSelection.Item(i)
            strOutput = strOutput & FormatEmailAsMarkdown(objMail)
            mailCount = mailCount + 1
        Else
            Debug.Print "Skipped item " & i & " (Type: " & TypeName(objSelection.Item(i)) & ")"
        End If
    Next i

    If mailCount = 0 Then
        MsgBox "No mail items found in selection.", vbExclamation
        GoTo Cleanup
    End If

    CopyToClipboard strOutput

    MsgBox mailCount & " email(s) copied to clipboard as Markdown."

Cleanup:
    Set objMail = Nothing
    Set objSelection = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ---------------------------------------------------------------
' Formats a single MailItem as Markdown with parsed reply chain
' ---------------------------------------------------------------
Private Function FormatEmailAsMarkdown(mail As MailItem) As String
    Dim strResult As String
    Dim strBody As String
    Dim segments As Variant
    Dim segCount As Long
    Dim j As Long

    ' Top-level heading: the email envelope
    strResult = "# " & mail.Subject & vbCrLf
    strResult = strResult & "**From:** " & mail.SenderName & _
                " | **Sent:** " & Format(mail.SentOn, "yyyy-mm-dd hh:nn") & vbCrLf
    If Len(mail.To) > 0 Then
        strResult = strResult & "**To:** " & mail.To & vbCrLf
    End If
    If Len(mail.CC) > 0 Then
        strResult = strResult & "**CC:** " & mail.CC & vbCrLf
    End If
    strResult = strResult & vbCrLf

    ' Get and clean the body
    strBody = mail.Body
    strBody = StripSignature(strBody)
    strBody = NormalizeWhitespace(strBody)

    ' Split into primary body + quoted replies
    segments = SplitReplyChain(strBody, segCount)

    ' Primary body gets ## heading
    If segCount >= 1 Then
        strResult = strResult & "## Body" & vbCrLf & vbCrLf
        strResult = strResult & NormalizeWhitespace(CStr(segments(0))) & vbCrLf & vbCrLf
    End If

    ' Quoted replies get ### headings
    For j = 1 To segCount - 1
        strResult = strResult & "### Quoted Reply " & j & vbCrLf & vbCrLf
        strResult = strResult & NormalizeWhitespace(CStr(segments(j))) & vbCrLf & vbCrLf
    Next j

    ' Trailing separator between top-level emails
    strResult = strResult & "---" & vbCrLf & vbCrLf

    FormatEmailAsMarkdown = strResult
End Function

' ---------------------------------------------------------------
' Splits email body at common reply/forward boundaries
' Returns Variant array of string segments
' ---------------------------------------------------------------
Private Function SplitReplyChain(ByVal strBody As String, ByRef segCount As Long) As Variant
    Dim markers(0 To 4) As String
    Dim positions() As Long
    Dim markerCount As Long
    Dim i As Long, j As Long
    Dim pos As Long
    Dim tempPos As Long
    Dim result() As String
    Dim prevChar As String
    Dim lookAhead As String

    ' Common reply/forward boundary patterns
    markers(0) = "-----Original Message-----"
    markers(1) = "________________________________"
    markers(2) = "From:"
    markers(3) = " wrote:"
    markers(4) = "-----Forwarded Message-----"

    ' Collect all boundary positions
    ReDim positions(0 To 0)
    markerCount = 0

    For i = 0 To UBound(markers)
        pos = 1
        Do
            tempPos = InStr(pos, strBody, markers(i), vbTextCompare)
            If tempPos = 0 Then Exit Do

            ' For "From:" — only count if it starts a line
            If markers(i) = "From:" Then
                If tempPos > 1 Then
                    prevChar = Mid(strBody, tempPos - 1, 1)
                    If prevChar <> vbLf And prevChar <> vbCr Then
                        pos = tempPos + 1
                        GoTo NextIteration
                    End If
                End If
                ' Verify this looks like a reply header block:
                ' Must have "Sent:" or "Date:" within the next 300 chars
                lookAhead = Mid(strBody, tempPos, SafeMin(300, Len(strBody) - tempPos + 1))
                If InStr(1, lookAhead, "Sent:", vbTextCompare) = 0 And _
                   InStr(1, lookAhead, "Date:", vbTextCompare) = 0 Then
                    pos = tempPos + 1
                    GoTo NextIteration
                End If
            End If

            ' Store position
            markerCount = markerCount + 1
            ReDim Preserve positions(0 To markerCount - 1)
            positions(markerCount - 1) = tempPos

            pos = tempPos + Len(markers(i))
NextIteration:
        Loop
    Next i

    ' If no boundaries found, return whole body as single segment
    If markerCount = 0 Then
        ReDim result(0 To 0)
        result(0) = strBody
        segCount = 1
        SplitReplyChain = result
        Exit Function
    End If

    ' Sort positions ascending (bubble sort — small array)
    Dim temp As Long
    For i = 0 To markerCount - 2
        For j = 0 To markerCount - 2 - i
            If positions(j) > positions(j + 1) Then
                temp = positions(j)
                positions(j) = positions(j + 1)
                positions(j + 1) = temp
            End If
        Next j
    Next i

    ' Remove near-duplicate positions (within 5 chars of each other)
    Dim uniquePos() As Long
    Dim uCount As Long
    ReDim uniquePos(0 To markerCount - 1)
    uniquePos(0) = positions(0)
    uCount = 1
    For i = 1 To markerCount - 1
        If positions(i) - uniquePos(uCount - 1) > 5 Then
            uniquePos(uCount) = positions(i)
            uCount = uCount + 1
        End If
    Next i
    ReDim Preserve uniquePos(0 To uCount - 1)

    ' Split body at boundary positions
    segCount = uCount + 1
    ReDim result(0 To segCount - 1)

    ' First segment: start to first boundary
    result(0) = Left(strBody, uniquePos(0) - 1)

    ' Middle segments
    For i = 0 To uCount - 2
        result(i + 1) = Mid(strBody, uniquePos(i), uniquePos(i + 1) - uniquePos(i))
    Next i

    ' Last segment
    result(uCount) = Mid(strBody, uniquePos(uCount - 1))

    SplitReplyChain = result
End Function

' ---------------------------------------------------------------
' Strips common email signature blocks
' ---------------------------------------------------------------
Private Function StripSignature(ByVal strBody As String) As String
    Dim sigMarkers(0 To 5) As String
    Dim i As Long
    Dim pos As Long
    Dim earliestPos As Long

    sigMarkers(0) = vbCrLf & "-- " & vbCrLf
    sigMarkers(1) = vbCrLf & "--" & vbCrLf
    sigMarkers(2) = "Sent from my iPhone"
    sigMarkers(3) = "Sent from my iPad"
    sigMarkers(4) = "Sent from Mail for Windows"
    sigMarkers(5) = "Get Outlook for"

    earliestPos = 0

    For i = 0 To UBound(sigMarkers)
        pos = InStr(1, strBody, sigMarkers(i), vbTextCompare)
        If pos > 0 Then
            ' Only strip if in the last 40% of the body
            If pos > Len(strBody) * 0.6 Then
                If earliestPos = 0 Or pos < earliestPos Then
                    earliestPos = pos
                End If
            End If
        End If
    Next i

    If earliestPos > 0 Then
        StripSignature = Left(strBody, earliestPos - 1)
    Else
        StripSignature = strBody
    End If
End Function

' ---------------------------------------------------------------
' Normalizes whitespace for token efficiency
' ---------------------------------------------------------------
Private Function NormalizeWhitespace(ByVal strText As String) As String
    Dim arrLines() As String
    Dim j As Long

    ' Collapse 3+ consecutive blank lines to 2
    Do While InStr(strText, vbCrLf & vbCrLf & vbCrLf) > 0
        strText = Replace(strText, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop

    ' Trim trailing spaces from each line
    arrLines = Split(strText, vbCrLf)
    For j = LBound(arrLines) To UBound(arrLines)
        arrLines(j) = Trim(arrLines(j))
    Next j
    strText = Join(arrLines, vbCrLf)

    ' Trim leading/trailing whitespace from the whole block
    strText = Trim(strText)

    NormalizeWhitespace = strText
End Function

' ---------------------------------------------------------------
' Safe Min function (avoids WorksheetFunction dependency)
' ---------------------------------------------------------------
Private Function SafeMin(a As Long, b As Long) As Long
    If a < b Then
        SafeMin = a
    Else
        SafeMin = b
    End If
End Function
