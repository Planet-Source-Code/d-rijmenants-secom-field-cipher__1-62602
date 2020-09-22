VERSION 5.00
Begin VB.Form frmCipher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SECOM Cipher  v1.0"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8415
   Icon            =   "frmCipher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   8175
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtGroupsPerLine 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "10"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblLenght 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   645
         Width           =   3015
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Key phrase"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Groups per line"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Menu mnuEncode 
      Caption         =   "&Encode"
   End
   Begin VB.Menu mnuDecode 
      Caption         =   "&Decode"
   End
   Begin VB.Menu mnuUndo 
      Caption         =   "&Undo"
   End
   Begin VB.Menu mnuClipBoard 
      Caption         =   "&To ClipBoard"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuPaperVersion 
      Caption         =   "Paper &Version"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strUndo As String

Private Sub Form_Load()
If Printer.DeviceName <> "" Then
   PrinterPresent = True
   Else
   PrinterPresent = False
   End If
End Sub

Private Sub mnuEncode_Click()
Dim tmpOutput As String
strUndo = Me.txtMain.Text
If Len(Me.txtMain.Text) > 10000 Then
    MsgBox "The text exceeds the limit of 10,000 characters.", vbCritical, " CS2C"
    Exit Sub
    End If
Screen.MousePointer = 11
tmpOutput = EncodeText(Me.txtMain.Text, Me.txtKey.Text)
If tmpOutput <> "" Then Me.txtMain.Text = MakeGroups(tmpOutput, True, Val(Me.txtGroupsPerLine.Text))
Screen.MousePointer = 0
End Sub

Private Sub mnuDecode_Click()
Dim tmpOutput As String
strUndo = Me.txtMain.Text
Screen.MousePointer = 11
tmpOutput = DecodeText(Me.txtMain.Text, Me.txtKey.Text)
If tmpOutput <> "" Then Me.txtMain.Text = tmpOutput
Screen.MousePointer = 0
End Sub

Private Sub mnuPaperVersion_Click()
loadPaperVersion ("SECOM")
frmPaperVersion.Show
End Sub

Private Sub mnuPrint_Click()
If Me.txtMain.Text = "" Then Exit Sub
If PrinterPresent = False Then Exit Sub
Call PrintString(Me.txtMain, 10, 10, 5, 5)
End Sub

Private Sub mnuUndo_Click()
Me.txtMain.Text = strUndo
End Sub

Private Sub mnuClipBoard_Click()
On Error Resume Next
Clipboard.SetText Me.txtMain.Text
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (vbModal)
End Sub

Private Sub txtGroupsPerLine_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtGroupsPerLine_GotFocus()
txtGroupsPerLine.SelStart = 0
txtGroupsPerLine.SelLength = Len(txtGroupsPerLine.Text)
End Sub

Private Sub txtKey_GotFocus()
txtKey.SelStart = 0
txtKey.SelLength = Len(txtKey.Text)
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
'returns only allowed characters
KeyAscii = UCase(KeyAscii)
If KeyAscii > 96 And KeyAscii < 123 Then
    'small letters, changed in capital
    KeyAscii = KeyAscii - 32
ElseIf (KeyAscii > 64 And KeyAscii < 91) Or KeyAscii = 32 Then
    'capital letters & spaces
    KeyAscii = KeyAscii
ElseIf KeyAscii < 32 Then
    'allow all specials
    KeyAscii = KeyAscii
Else
    'don't use
    KeyAscii = 0
End If
End Sub

Private Sub txtMain_Change()
Me.lblLenght.Caption = Str(Len(Me.txtMain.Text)) & " Chars"
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
If Len(Me.txtMain.Text) > 9999 And KeyAscii <> 8 Then
    MsgBox "Text limited to 10,000 characters.", vbCritical
    KeyAscii = 0
    End If
'returns only allowed characters
If KeyAscii > 96 And KeyAscii < 123 Then
    'small letters, changed in capital
    KeyAscii = KeyAscii - 32
ElseIf KeyAscii > 64 And KeyAscii < 91 Then
    'capital letters
    KeyAscii = KeyAscii
ElseIf KeyAscii > 47 And KeyAscii < 58 Then
    'numbers
    KeyAscii = KeyAscii
ElseIf KeyAscii = 32 Then
    'spaces
    KeyAscii = KeyAscii
ElseIf KeyAscii < 32 Then
    'allow all specials
    KeyAscii = KeyAscii
Else
    'don't use
    KeyAscii = 0
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmPaperVersion
Unload frmAbout
End Sub

