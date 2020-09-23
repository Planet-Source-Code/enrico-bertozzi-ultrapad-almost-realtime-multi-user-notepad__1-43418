VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form cMain 
   Caption         =   "Quad Edit"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10605
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update (F5)"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   7575
   End
   Begin MSWinsockLib.Winsock wClientData 
      Left            =   9360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   51334
   End
   Begin MSWinsockLib.Winsock wClient 
      Left            =   9360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   51333
   End
   Begin VB.TextBox tChat 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8055
   End
   Begin VB.TextBox vChat 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox tMain 
      Height          =   4575
      Left            =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox tEdit 
      Height          =   2415
      Left            =   2880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5640
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   0   'False
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0EC2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lUsers 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   5
      Top             =   165
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "chat:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   165
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnsConnect 
         Caption         =   "Connect to..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnsClose 
         Caption         =   "Close session"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnf1 
         Caption         =   "-"
      End
      Begin VB.Menu mnfExit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "cMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub reset()
'reset variables and close connections
mainenabled = False
Me.Caption = Version
wClient.Close
wClientData.Close
tEdit.text = ""
lUsers.Caption = ""
End Sub

Private Sub cmdUpdate_Click()
If Not mainenabled Then Exit Sub
'If user hasn't changed edit txtbox, don't send anything to avoid useless traffic
If Not editchanged Then Exit Sub
wClientData.SendData myId & vbCrLf & tEdit.text
'put many doevents if things doesn't work! no, really, it may help
DoEvents

editchanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'hand-made shortcut
If KeyCode = vbKeyF5 Then cmdUpdate_Click
End Sub

Private Sub Form_Load()
reset

ReDim ptrs(0) As Long
ReDim ptrl(0) As Long

'define syntax-highlighting words
word(0) = "return"
word(1) = "int"
word(2) = "long"
word(3) = "float"
word(4) = "double"
word(5) = "do"
word(6) = "while"
word(7) = "if"
word(8) = "char"
word(9) = "void"
word(10) = "sizeof"
word(11) = "typedef"
word(12) = "#include"
word(13) = "#define"
word(14) = "switch"
word(15) = "case"
word(16) = "break"
word(17) = "continue"
word(18) = "else"
word(19) = "default"
word(20) = "for"
word(21) = "signed"
word(22) = "unsigned"
word(23) = "class"
word(24) = "public"
word(25) = "private"
word(26) = "protected"
word(27) = "struct"
word(28) = "bool"
word(29) = "#pragma"
word(30) = "#ifdef"
word(31) = "#endif"
End Sub

Private Sub Form_Resize()
'if form is minimized, do nothing
If Me.WindowState = vbMinimized Then Exit Sub

'This manages the form layout. If window is resized too small, avoid applying layout.
If ScaleHeight < 400 Then
    Me.Height = h * Screen.TwipsPerPixelY
    Exit Sub
End If

If ScaleWidth < 400 Then
    Me.Width = w * Screen.TwipsPerPixelX
    Exit Sub
End If

vChat.Left = 0
vChat.Width = 200
tMain.Left = 201
tMain.Width = ScaleWidth - 200
tEdit.Width = ScaleWidth - 200
tEdit.Left = 201

vChat.Top = 32
tMain.Top = 32
tEdit.Top = ScaleHeight - 275
cmdUpdate.Top = ScaleHeight - 300
cmdUpdate.Height = 25
cmdUpdate.Left = 201
cmdUpdate.Width = ScaleWidth - 200
vChat.Height = ScaleHeight - vChat.Top
tMain.Height = ScaleHeight - tMain.Top - 300
tEdit.Height = 275
tChat.Width = ScaleWidth - tChat.Left - 130
lUsers.Left = ScaleWidth - 120

w = ScaleWidth
h = ScaleHeight
End Sub

Private Sub mnsClose_Click()
reset
End Sub

Private Sub mnsConnect_Click()
Dim t As String

'ask the user for IP or name server
t = InputBox("IP address or DNS:")
'and exit if he has entered none
If Len(t) = 0 Then Exit Sub

'connect the command and data socket
wClient.Connect t
wClientData.Connect t, 51334
'show the wait window
ShowWait "waiting (connecting...)"
End Sub

Private Sub tChat_KeyDown(KeyCode As Integer, Shift As Integer)
'if users press Enter, send text
'text will be returned from server
If KeyCode = vbKeyReturn And mainenabled Then
    Send "C" & tChat.text & vbCrLf
    tChat.text = ""
End If
End Sub

Private Sub tEdit_Change()
editchanged = True
End Sub

Private Sub tMain_Change()
Dim pos As Long, pos2 As Long, pos3 As Long, I As Long
Dim comment As Boolean
Dim bpos As Long, blen As Long, endparse As Boolean

bpos = tMain.SelStart
blen = tMain.SelLength

'Syntax highlighter

'first of all, set all text to plain black
tMain.SelStart = 0
tMain.SelLength = Len(tMain.text)
tMain.SelBold = False
tMain.SelColor = 0

'then, use the Find method to find each word, and hilight it in blue bold
For I = 0 To UBound(word)
    pos = 0
    Do
        'start searching from POS
        pos = tMain.Find(word(I), pos, , 14)
        If pos = -1 Then Exit Do
        
        tMain.SelStart = pos
        tMain.SelLength = Len(word(I))
        tMain.SelBold = True
        tMain.SelColor = 16711680

        'update pos
        pos = pos + Len(word(I))
    Loop
Next I


'search for the comment '//' symbols
pos3 = 0
Do
    endparse = True
    pos = tMain.Find("//", pos3, , 8)
    If pos <> -1 Then
        'if found, search for the next carriage return
        pos2 = tMain.Find(vbCrLf, pos, , 8)
        tMain.SelStart = pos
        'if carriage return is not found, we've reached the end of file
        If pos2 = -1 Then tMain.SelLength = Len(tMain.text) - pos Else tMain.SelLength = pos2 - pos
        tMain.SelBold = False
        tMain.SelColor = 32768
        pos3 = pos2
        endparse = False
    End If
Loop Until endparse

'search for the comment '/* ... */' symbols
pos3 = 0
Do
    endparse = True
    'search for the first symbol
    pos = tMain.Find("/*", pos3, , 8)
    If pos <> -1 Then
        'if found, search for the second symbol, starting from the first sym's place
        pos2 = tMain.Find("*/", pos, , 8)
        If pos2 <> -1 Then
            tMain.SelStart = pos
            tMain.SelLength = (pos2 - pos) + 2
            tMain.SelBold = False
            tMain.SelColor = 32768
            pos3 = pos2
            endparse = False
        End If
    End If
Loop Until endparse

'see below
Underline

tMain.SelStart = bpos
tMain.SelLength = blen
End Sub

Private Sub Underline()
Exit Sub '<------------- not enabled
'I would enable that, but it causes too much flickering
'it consisted on underlining what other users were editing

Dim I As Long

tMain.SelStart = 0
tMain.SelLength = Len(tMain.text)
tMain.SelUnderline = False

For I = 0 To nClients
    If I <> myId Then
        tMain.SelStart = ptrs(I)
        tMain.SelLength = ptrl(I)
        tMain.SelUnderline = True
    End If
Next I

tMain.SelStart = 0

End Sub

Private Sub tMain_Click()
'To determine pointer movements, I use Click and KeyUp events instead of SelChange
'because SelChange would be called a lot of times uselessly on syntax hilighter
checkUpd
tEdit.SetFocus
End Sub

Private Sub checkUpd()
Dim I As Long
If Not mainenabled Then Exit Sub

'If user hasn't pressed F5, remind it
If editchanged Then
    If MsgBox("Update file first?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        cmdUpdate_Click
        Exit Sub
    End If
End If

'negotiate a new position. read below for commands
If tMain.SelStart > 100 Then
    Send "P" & myId & "," & tMain.SelStart - 100 & vbCrLf
Else
    Send "P" & myId & ",1" & vbCrLf
End If
Send "L" & myId & ",200" & vbCrLf
End Sub

Private Sub tMain_KeyUp(KeyCode As Integer, Shift As Integer)
checkUpd
tEdit.SetFocus
End Sub

Private Sub wClient_Close()
MsgBox "Disconnected from server", vbCritical
reset
End Sub

Private Sub wClient_Connect()
'update wait window
WaitText "waiting (connected)"
End Sub

Private Sub wClient_DataArrival(ByVal bytesTotal As Long)
Dim cmd As String, cmdp As String
Dim ncmd As Long, I As Long, idx As Long, temp As Long, D As Long

'get the commandline
wClient.GetData cmd, vbString

'get the number of commands
ncmd = getParams(cmd)

For I = 1 To ncmd
    'get the command
    cmdp = strParam(cmd, I)
    Select Case Left(cmdp, 1)
        'P: set user's pointer start to pos
        'P<user>,<pos>
        Case "P"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        temp = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        ptrs(idx) = temp
        ptrl(idx) = 0
        'update debug labels
        DebugLabels

        'L: set user's pointer len to value
        'L<user>,<value>
        Case "L"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        temp = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        ptrl(idx) = temp
        'if I am the user, i set the txtbox text to the interval
        If idx = myId Then
            tEdit.text = Mid(tMain.text, ptrs(myId), ptrl(myId))
            editchanged = False
            tEdit.Locked = False
        End If
        Underline
        DebugLabels

        Case "C"
        'update chat
        vChat.text = Mid(cmdp, 2) & vbCrLf & "__________________" & vbCrLf & vbCrLf & vChat.text

        'K: set user's pointer start to pos (silent)
        'used to change pointer without resetting lenght
        'if an user whose area is below that current user's one, this command
        'is used instead of P
        Case "K"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        ptrs(idx) = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        Underline
        DebugLabels

        'not currently used: set user's pointer lenght
        Case "W"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        ptrl(idx) = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        DebugLabels

        'assign user ID and activate txtboxs
        'A<id>,<nclients>
        Case "A"
        myId = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        nClients = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        lUsers.Caption = nClients & " user(s) connected"
        tEdit.Enabled = True
        tChat.Enabled = True
        ReDim ptrs(nClients) As Long
        ReDim ptrl(nClients) As Long
        mainenabled = True
        Load cDebug
        For D = 1 To nClients - 1
            Load cDebug.lDebug(I)
            cDebug.lDebug(D).Top = 24 * I
            cDebug.lDebug(D).Left = 0
            cDebug.lDebug(D).Width = 289
            cDebug.lDebug(D).Height = 17
            cDebug.lDebug(D).Visible = True
        Next D
        cDebug.lDebug(myId).FontBold = True
        cDebug.lMyid.Caption = "myId = " & myId
        cDebug.Show , Me
        Underline
        HideWait
        
        'refuse user's negotiation and emit beep
        Case "X"
        Beep
        tEdit.Locked = True
        tEdit.text = ""
    End Select
Next I
End Sub

Private Sub wClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
HideWait
MsgBox "Error during connection", vbExclamation
End Sub

Private Sub wClientData_Close()
MsgBox "Data socket disconnected from server", vbCritical
reset
End Sub

Private Sub wClientData_DataArrival(ByVal bytesTotal As Long)
Dim id As Long, buftext As String, temp As String, hdt As Long
wClientData.GetData temp, vbString

hdt = InStr(1, temp, vbCrLf)
'get message sender
id = CInt(Left(temp, hdt))
'temp=message
temp = Mid(temp, hdt + Len(vbCrLf))

'if ID=-1 the message is a file being displayed on main txtbox
If id = -1 Then
    tMain.text = temp
Else
    'else, replace the current text
    buftext = tMain.text
    buftext = Left(buftext, ptrs(id) - 1) & temp & Mid(buftext, ptrs(id) + ptrl(id))
    tMain.text = buftext
End If
End Sub

Private Sub wClientData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Unable to connect data socket to the server", vbExclamation
reset
End Sub

Private Sub DebugLabels()
Dim I As Integer

For I = 0 To nClients - 1
    cDebug.lDebug(I).Caption = "START " & ptrs(I) & ", LEN " & ptrl(I)
Next I
End Sub
