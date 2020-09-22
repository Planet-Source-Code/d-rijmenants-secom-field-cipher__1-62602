Attribute VB_Name = "modGeneral"
'----------------------------------------------------------------
'
'                           General Subs
'
'----------------------------------------------------------------

Option Explicit

Public Function MakeGroups(TextIn As String, Groups As Boolean, GroupsPerLine As Integer) As String
'make groups of 5 characters each
Dim i As Long
If Groups = False Or GroupsPerLine = 0 Then MakeGroups = TextIn: Exit Function
For i = 1 To Len(TextIn)
    MakeGroups = MakeGroups & Mid(TextIn, i, 1)
    If i Mod 5 = 0 Then MakeGroups = MakeGroups & " "
    If i Mod (GroupsPerLine * 5) = 0 Then MakeGroups = MakeGroups & vbCrLf
Next
End Function

Public Sub loadPaperVersion(aTitle As String)
'load the help on the pencil-and-paper version
Dim FileO As Integer
Dim strInput As String
On Error GoTo errHandler
FileO = FreeFile
Open App.Path & "\" & aTitle & ".txt" For Input As #FileO
strInput = Input(LOF(FileO), 1)
Close FileO
frmPaperVersion.Text1.Text = strInput
Exit Sub
errHandler:
Close FileO
End Sub




