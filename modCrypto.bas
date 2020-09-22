Attribute VB_Name = "modCrypto"
'----------------------------------------------------------------
'
'                     SECOM Encryption subs
'
'----------------------------------------------------------------

Option Explicit

'1st transposition
Public W1 As Byte           'width
Public T1(20) As Byte       'key
Public Tbox1() As Byte      'array

'2nd transposition
Public W2 As Byte           'width
Public T2(20) As Byte       'key
Public Tbox2() As Byte      'array
Public Triangles() As Byte  'array holding trangular areas

'checkerboard
Public TS(10) As Byte       'key
Public TSS(40) As String    'key with Nr's as string
Public CheckerBoard(40) As String 'array of combinations fi "A27"

Public TotalDigits As Long

Public Function EncodeText(ByVal aText As String, ByVal aKey As String)
Dim K           As Long
Dim J           As Long
Dim Count       As Long
Dim Pos         As Long
Dim T1out()     As Byte
Dim Sout()      As Byte
Dim Col         As Byte
Dim tmp         As String
Dim Sign        As String
Dim FoundFlag   As Boolean
Dim WarnFlag    As Boolean

If Len(aText) < 1 Then
    MsgBox "Please enter a message to encode.", vbCritical
    Exit Function
    End If

'delete spaces in key phrase
For K = 1 To Len(aKey)
    If Mid(aKey, K, 1) <> " " Then tmp = tmp & Mid(aKey, K, 1)
Next K
aKey = tmp

If Len(aKey) < 20 Then
    MsgBox "Please enter a key phrase with at least 20 letters.", vbCritical
    Exit Function
    End If

aKey = Left(aKey, 20)

Call InitKey(aKey)

'checkerboard, encode plain text to Sout()
Count = 0
WarnFlag = False
For K = 1 To Len(aText)
    Sign = UCase(Mid(aText, K, 1))
    'skip lf and tabs
    If Sign = Chr(8) Or Sign = Chr(9) Or Sign = Chr(10) Then WarnFlag = True: GoTo skipSigns
    'replace linefeed with space
    If Sign = Chr(13) Then
        WarnFlag = True
        If K + 1 <= Len(aText) Then
            If UCase(Mid(aText, K + 1, 1)) = Chr(10) Then
                Sign = " "
                K = K + 1
            End If
        End If
    End If
    FoundFlag = False
    For J = 1 To 40
        If Sign = Left(CheckerBoard(J), 1) Then
            'found!
            FoundFlag = True
            'search the sign at the checkerboard
            If Len(CheckerBoard(J)) = 2 Then
                'one figure
                Count = Count + 1
                ReDim Preserve Sout(Count)
                Sout(Count) = Val(Right(CheckerBoard(J), 1))
                Else
                'two figures
                Count = Count + 2
                ReDim Preserve Sout(Count)
                Sout(Count - 1) = Val(Mid(CheckerBoard(J), 2, 1))
                Sout(Count) = Val(Right(CheckerBoard(J), 1))
            End If
        End If
    Next J
    If FoundFlag = False Then WarnFlag = True
        
skipSigns:
Next K

If WarnFlag = True Then MsgBox "The message contains other characters than letters, numbers or spaces." & vbCrLf & vbCrLf & "The unallowed characters were skipped during encryption.", vbExclamation

'make sure we have complete groups of 5 at the end (add zero's)
K = Count Mod 5
If K <> 0 Then
    K = 5 - K
    Count = Count + K
    ReDim Preserve Sout(Count)
    For J = (Count - K) + 1 To Count
        Sout(J) = 0
    Next
    End If


TotalDigits = Count
ReDim Tbox1(TotalDigits) As Byte
ReDim Tbox2(TotalDigits) As Byte
ReDim T1out(TotalDigits) As Byte


'set triangles matrix of disrupted transposition
Call SetTriangles

'transfer checkerboard conversion to 1st transposition tbox1()
For K = 1 To TotalDigits
    Tbox1(K) = Sout(K)
Next K

'read off tbox1(), col by col, to t1out()
Count = 1
For K = 1 To W1
    Col = T1(K)
    For J = 0 + Col To TotalDigits Step W1
        T1out(Count) = Tbox1(J)
        Count = Count + 1
    Next
Next

'transfer t1out() to 2nd transposition tbox2

'first avoiding triangle areas, according to matrix
Count = 1
For K = 1 To TotalDigits
    'first add according non marked areas
    If Triangles(K) <> 255 Then
        Tbox2(K) = T1out(Count)
        Count = Count + 1
    End If
Next

'fill triangles
For K = 1 To TotalDigits
    'add only according marked areas
    If Triangles(K) = 255 Then
        Tbox2(K) = T1out(Count)
        Count = Count + 1
    End If
Next K

'read off tbox2, col by col
For K = 1 To W2
    Col = T2(K)
    For J = 0 + Col To TotalDigits Step W2
        'transfer to text
        EncodeText = EncodeText & Trim(Str(Tbox2(J)))
    Next
Next

End Function

Public Function DecodeText(ByVal aText As String, ByVal aKey As String)
Dim K           As Long
Dim J           As Long
Dim Count       As Long
Dim Pos         As Long
Dim Col         As Byte
Dim Tasc        As Byte
Dim TextIn()    As Byte
Dim T1In()      As Byte
Dim Sin()       As Byte
Dim Sval        As Byte
Dim tmp         As String
Dim Sign        As String
Dim FoundFlag   As Boolean

If Len(aText) < 1 Then
    MsgBox "Please enter number groups to decode.", vbCritical
    Exit Function
    End If

    
For K = 1 To Len(aKey)
    If Mid(aKey, K, 1) <> " " Then tmp = tmp & Mid(aKey, K, 1)
Next K
aKey = tmp

If Len(aKey) < 20 Then
    MsgBox "Please enter a key phrase with at least 20 letters.", vbCritical
    Exit Function
End If

aKey = Left(aKey, 20)

'put code text in TextIn(), only taking numbers
Count = 0
For K = 1 To Len(aText)
    Tasc = Asc(Mid(aText, K, 1))
    If Tasc > 47 And Tasc < 58 Then
        Count = Count + 1
        ReDim Preserve TextIn(Count) As Byte
        TextIn(Count) = Tasc - 48
        End If
Next K

TotalDigits = Count
ReDim Tbox1(TotalDigits) As Byte
ReDim Tbox2(TotalDigits) As Byte
ReDim T1In(TotalDigits) As Byte
ReDim Sin(TotalDigits) As Byte

Call InitKey(aKey)
Call SetTriangles

'fill 2nd transposition tbox2() col by col
Count = 1
For K = 1 To W2
    Col = T2(K)
    For J = 0 + Col To TotalDigits Step W2
        Tbox2(J) = TextIn(Count)
        Count = Count + 1
    Next
Next

'read off tbox2() to 1st transposition T1In(), row by row
'first avoiding triangle areas
Count = 1
For K = 1 To TotalDigits
    If Triangles(K) <> 255 Then
        T1In(Count) = Tbox2(K)
        Count = Count + 1
    End If
Next
' now read triangles
For K = 1 To TotalDigits
    If Triangles(K) = 255 Then
        T1In(Count) = Tbox2(K)
        Count = Count + 1
    End If
Next K

'fill 1st transposition with T1In, col by col
Count = 1
For K = 1 To W1
    Col = T1(K)
    For J = 0 + Col To TotalDigits Step W1
        Tbox1(J) = T1In(Count)
        Count = Count + 1
    Next
Next

'read off 1st transposition, row by row, and decode with checkerboard
Count = 0
For K = 1 To TotalDigits
    Sval = Tbox1(K)
    If Sval = TS(3) Or Sval = TS(6) Or Sval = TS(9) Then
        'two number code combinations
        K = K + 1
        If K > TotalDigits Then 'error last digit overflow combination
            DecodeText = DecodeText & "?"
            Exit Function
            End If
        Sign = Trim(Str(Sval)) & Trim(Str(Tbox1(K)))
        FoundFlag = False
        'go through all 2 digit combinations
        For J = 8 To 40
            'search for sign combination in CheckerBoard
            If Mid(CheckerBoard(J), 2) = Sign Then
                DecodeText = DecodeText & Left(CheckerBoard(J), 1)
                FoundFlag = True
                Exit For
            End If
        Next J
        If FoundFlag = False Then
            DecodeText = DecodeText & "?"
            End If
        Else
        'one number code combinations
        Sign = Trim(Str(Sval))
        FoundFlag = False
        'go through all 1 digit combinations
        For J = 1 To 7
            'search for sign combination in CheckerBoard
            If Right(CheckerBoard(J), 1) = Sign Then
                DecodeText = DecodeText & Left(CheckerBoard(J), 1)
                FoundFlag = True
                Exit For
            End If
        Next J
        If FoundFlag = False Then
            DecodeText = DecodeText & "?"
            End If
        End If
Next K

End Function


Public Sub InitKey(ByVal aKey As String)
Dim K           As Long
Dim J           As Long
Dim Count       As Long
Dim tmpNr(20)   As Byte
Dim tmpLFG(10)  As Byte
Dim K1(10)      As Byte
Dim K2(10)      As Byte
Dim T12(50)     As Byte
Dim LFG(60)     As Byte
Dim Z           As Byte
Dim Col         As Byte
Dim Smallest    As Byte
Dim Spos        As Integer
Dim KS1         As String
Dim KS2         As String
Dim UsedNr      As String
Dim ZS          As String
Dim Crow        As String
Dim Flag        As Boolean

KS1 = Left(aKey, 10)
KS2 = Right(aKey, 10)

'transfer letters 1st half of key to temp array
For K = 1 To 10
    tmpNr(K) = Asc(Mid(KS1, K, 1))
Next

'number 1st half of key to K1()
For K = 1 To 10
    Z = 255
    For J = 1 To 10
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    K1(Smallest) = K
Next

'transfer letters 2nd half of key to temp array
For K = 1 To 10
    tmpNr(K) = Asc(Mid(KS2, K, 1))
Next

'number 1st half of key to K1()
For K = 1 To 10
    Z = 255
    For J = 1 To 10
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    K2(Smallest) = K
Next

'initialise LFG random number generator
For K = 1 To 10
    Z = K1(K) + K2(K)
    If Z > 9 Then Z = Z - 10
    LFG(K) = Z
Next

'generate 50 digits in LFG
For K = 1 To 50
    Z = LFG(K) + LFG(K + 1)
    If Z > 10 Then Z = Z - 10
    LFG(K + 10) = Z
Next K

'shift all rows to begin LFG
For K = 1 To 50
    LFG(K) = LFG(K + 10)
    'hold last row for checkerboard nummering
    If K > 40 Then tmpNr(K - 40) = LFG(K)
Next K

'number row for checkerboard
For K = 1 To 10
    Z = 255
    For J = 1 To 10
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    TS(Smallest) = K
    
Next

'create checkerboard combinations in CheckerBoard() fi "F53" etc...
'create string version of digits
For K = 1 To 10
    TSS(K) = Trim(Str(TS(K)))
    If TSS(K) = "10" Then TSS(K) = "0"
Next K
'create first row, 1 digit combination
Crow = "ESTONIA"
CheckerBoard(1) = Mid(Crow, 1, 1) & TSS(1)
CheckerBoard(2) = Mid(Crow, 2, 1) & TSS(2)
CheckerBoard(3) = Mid(Crow, 3, 1) & TSS(4)
CheckerBoard(4) = Mid(Crow, 4, 1) & TSS(5)
CheckerBoard(5) = Mid(Crow, 5, 1) & TSS(7)
CheckerBoard(6) = Mid(Crow, 6, 1) & TSS(8)
CheckerBoard(7) = Mid(Crow, 7, 1) & TSS(10)

'createsecond row, shifted
Crow = "BCDFGHJKLM"
For K = 1 To 10
    Spos = K - (TS(3) - 1)
    If Spos < 1 Then Spos = Spos + 10
    CheckerBoard(7 + K) = Mid(Crow, Spos, 1) & TSS(3) & TSS(K)
Next K

'create thirth row, shifted
Crow = "PQRUVWXYZ "
For K = 1 To 10
    Spos = K - (TS(6) - 1)
    If Spos < 1 Then Spos = Spos + 10
    CheckerBoard(17 + K) = Mid(Crow, Spos, 1) & TSS(6) & TSS(K)
Next K

'create fourth row (shifted)
Crow = "1234567890"
For K = 1 To 10
    Spos = K - (TS(9) - 1)
    If Spos < 1 Then Spos = Spos + 10
    CheckerBoard(27 + K) = Mid(Crow, Spos, 1) & TSS(9) & TSS(K)
Next K
'now change 10 in TS() into 0
For K = 1 To 10
    If TS(K) = 10 Then TS(K) = 0
Next K

'add CheckerBoard numbers TS with 2nd half key numbers K2 for LFG transposition
For K = 1 To 10
    Z = K2(K) + TS(K)
    If Z > 10 Then Z = Z - 10
    tmpNr(K) = Z
Next K

'number result of adding TS and T2 into tmpLFG()
For K = 1 To 10
    Z = 255
    For J = 1 To 10
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    tmpLFG(K) = Smallest
Next

'read of width for 1st transposition t1 and 2nd transposition t2
W1 = 0
W2 = 0
UsedNr = ""
Flag = False

For K = 50 To 1 Step -1
    Z = LFG(K)
    If Z = 10 Then Z = 0 'tread 10's as 0
    ZS = Trim(Str(Z))
    If InStr(1, UsedNr, ZS) = False Then
        UsedNr = UsedNr + ZS
        If Flag = False Then
            W1 = W1 + Z
            If W1 > 9 Then Flag = True: UsedNr = ""
        Else
            W2 = W2 + Z
            If W2 > 9 Then Exit For
        End If
    End If
Next

'read off digit from LFG, col by col according tmplfg(x) into T12(x)
Count = 1
For K = 1 To 10
    Col = tmpLFG(K)
    For J = 0 + Col To 50 Step 10
        T12(Count) = LFG(J)
        Count = Count + 1
    Next
Next

'numbering T1
For K = 1 To W1
    tmpNr(K) = T12(K)
Next K

For K = 1 To W1
    Z = 255
    For J = 1 To W1
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    T1(K) = Smallest
Next

'numbering T2
For K = 1 To W2
    tmpNr(K) = T12(K + W1)
Next K

For K = 1 To W2
    Z = 255
    For J = 1 To W2
        If tmpNr(J) < Z Then
            Z = tmpNr(J)
            Smallest = J
        End If
    Next
    tmpNr(Smallest) = 255
    T2(K) = Smallest
Next

End Sub

Public Sub SetTriangles()
'set disrupted trangles
Dim Tseq(20) As Byte
Dim Cmax As Byte
Dim K As Long
Dim J As Long
Dim Pos As Long
Dim Count As Long

'set sequence columns
For K = 1 To W2
    Tseq(K) = T2(K)
Next K

'draw trangle areas as 255
ReDim Triangles(TotalDigits) As Byte
Count = 1
Cmax = Tseq(Count)
For K = 1 To TotalDigits Step W2
For J = 1 To W2
    If J >= Cmax Then
        'draw rest of line as 255
        Pos = (K + (J - 1))
        If Pos > TotalDigits Then Exit Sub
        Triangles(Pos) = 255
    End If
Next J
Cmax = Cmax + 1
If Cmax > W2 Then
    'this triangle finished, make next one
    Count = Count + 1
    If Count > W2 Then Exit Sub
    Cmax = Tseq(Count)
    K = K + W2
    End If
Next

End Sub
