Attribute VB_Name = "modPrint"
'----------------------------------------------------------------
'
'  Print module v3 (c) D.Rijmenants
'  PrintString (Text, leftfmargin, rightmargin, topmargin, bottommargin)
'  margins are long values 0-100 percent
'
'----------------------------------------------------------------

Public PrinterPresent As Boolean
Option Explicit

Public Sub PrintString(printVar As String, leftMargePrcnt As Long, rightMargePrcnt As Long, topMargePrcnt As Long, bottomMargePrcnt As Long)
Dim lMarge As Long
Dim rMarge As Long
Dim tMarge As Long
Dim bMarge As Long
Dim printLijn As String
Dim staPos  As Long
Dim endPos As Long
Dim txtHoogte As Long
Dim printHoogte As Long
Dim objectHoogte As Long
Dim objectBreedte As Long
Dim currYpos As Long
Dim cutChar As String
Dim K As Long
Dim cutPos As Long

On Error Resume Next

Screen.MousePointer = 11

Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False
Printer.FontStrikethru = False

txtHoogte = Printer.TextHeight("AbgWq")
lMarge = Int((Printer.Width / 100) * leftMargePrcnt)
rMarge = Int((Printer.Width / 100) * rightMargePrcnt)
tMarge = Int((Printer.Height / 100) * topMargePrcnt)
bMarge = Int((Printer.Height / 100) * bottomMargePrcnt)
objectHoogte = Printer.Height - tMarge - bMarge
objectBreedte = Printer.Width - lMarge - rMarge
Printer.CurrentY = tMarge
staPos = 1
endPos = 0
Do

'get next line to crlf
endPos = InStr(staPos, printVar, vbCrLf)
If endPos <> 0 Then
    printLijn = Mid(printVar, staPos, endPos - staPos)
    Else
    printLijn = Mid(printVar, staPos)
    endPos = Len(printVar)
    End If
    
'check lenght one line
If Printer.TextWidth(printLijn) <= objectBreedte Then
    'line ok, keep line as it is
    staPos = endPos + 2
    Else
    'line to big, try to cut of at space or other signs within limits
    cutPos = 0
    For K = 1 To Len(printLijn)
        cutChar = Mid(printLijn, K, 1)
        If cutChar = " " Or cutChar = "." Or cutChar = "," Or cutChar = ":" Or cutChar = ")" Then
            If Printer.TextWidth(Left(printLijn, K)) > objectBreedte Then Exit For
            cutPos = K
        End If
    Next K
    'check result search for space
    If cutPos > 1 Then
        'cut off on space
        printLijn = Mid(printVar, staPos, cutPos)
        staPos = staPos + cutPos
        Else
        'no cut-character found within limits, so cut line on paperwidth
        For K = 1 To Len(printLijn)
            If Printer.TextWidth(Left(printLijn, K)) > objectBreedte Then Exit For
        Next K
        printLijn = Mid(printVar, staPos, K - 1)
        staPos = staPos + (K - 1)
    End If
End If
'print line
Printer.CurrentX = lMarge
currYpos = Printer.CurrentY + txtHoogte
If currYpos > (tMarge + objectHoogte) - txtHoogte Then
    Printer.NewPage
    Printer.CurrentY = tMarge
    Printer.CurrentX = lMarge
    End If
Printer.Print printLijn
Loop While staPos < Len(printVar)
Printer.EndDoc
Screen.MousePointer = 0
End Sub


