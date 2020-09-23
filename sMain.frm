VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form sMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UltraPad Server 3.0"
   ClientHeight    =   975
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5325
   Icon            =   "sMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdfs 
      Left            =   2760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrUpd 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   600
   End
   Begin MSWinsockLib.Winsock wSD 
      Index           =   0
      Left            =   4800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   51334
   End
   Begin MSWinsockLib.Winsock wSM 
      Index           =   0
      Left            =   4800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   51333
   End
   Begin MSWinsockLib.Winsock wServer 
      Left            =   3600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   51333
   End
   Begin VB.Label lUsers 
      Caption         =   "0"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lRecv 
      Caption         =   "0 bytes"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lSent 
      Caption         =   "0 bytes"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Received data:"
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
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Sent data:"
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
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Users connected:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New session"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save as"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "sMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'set main variables
Me.Caption = Version
WaitIncoming = False
numSockets = 1
datain = 0
dataout = 0
FileSaved = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim R As Long

'ask the user if "he is sure" blah blah
If Not FileSaved Then
    R = MsgBox("Save the file first?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
    If R = vbYes Then
        mnuSave_Click
    ElseIf R = vbCancel Then Cancel = True
    End If
End If
End Sub

Private Sub mnuNew_Click()
Dim t As String

'ask the user how many user can connect to the server
t = InputBox("Users to allow:")
'exit if he types nothing
If Len(t) = 0 Then Exit Sub

nClients = CLng(t)
WaitIncoming = True 'waiting for incoming connections
FileSaved = True
progr = 0 'set progressive ID count to zero

sDebug.Show , Me

lUsers.Caption = progr & "/" & nClients & " (waiting for connections...)"
wServer.Listen
End Sub

Private Sub mnuSave_Click()
'show the dialog
cdfs.DialogTitle = "Save as"
cdfs.Filter = "All files|*.*"
cdfs.Flags = cdlOFNOverwritePrompt 'tell the dialog that we want to be warned
                                    'if the file may be overwritten
cdfs.ShowSave

If Len(cdfs.FileName) = 0 Then Exit Sub

'save the file
SaveFile cdfs.FileTitle, alltext
FileSaved = True
End Sub

Private Sub tmrUpd_Timer()
'update the stats
lRecv.Caption = datain & " bytes"
lSent.Caption = dataout & " bytes"
End Sub

Private Sub wSD_Close(Index As Integer)
'as said before, if one user disconnects, the server disconnects all the other
'users

Dim I As Long

For I = 0 To nClients - 1
    wSD(I).Close
    wSM(I).Close
Next I
End Sub

Private Sub wSD_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'accept  any incoming data connection
wSD(Index).Close
wSD(Index).Accept requestID
End Sub

Private Sub wSD_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'process data we eventually receive from data socket

Dim I As Long, id As Long, hdt As Long
Dim temp As String

'update variable
datain = datain + bytesTotal
wSD(Index).GetData temp, vbString

'split data and header
hdt = InStr(1, temp, vbCrLf)
id = CInt(Left(temp, hdt))
'temp=data
temp = Mid(temp, hdt + Len(vbCrLf))

'update the file with data coming from user
alltext = Left(alltext, ptrs(id) - 1) & temp & Mid(alltext, ptrs(id) + ptrl(id))

'here we check if there are any user assigned to areas greater to the sender user.
'if there are, they have to be notified of a change in their own area because
'of the new data. this only uses the K command
For I = 0 To nClients - 1
    If ptrs(I) > ptrs(id) + ptrl(id) Then
        SendAll "K" & I & "," & (ptrs(I) + (Len(temp) - ptrl(id))) & vbCrLf
        ptrs(I) = ptrs(I) + (Len(temp) - ptrl(id))
    End If
Next I
DebugLabels

'ping data to all users
For I = 0 To nClients - 1
    wSD(I).SendData id & vbCrLf & temp
    DoEvents
Next I

FileSaved = False
End Sub

Private Sub wSD_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'if send is complete, update counter
If bytesRemaining = 0 Then dataout = dataout + bytesSent
End Sub

Private Sub wServer_ConnectionRequest(ByVal requestID As Long)
Dim created As Boolean, I As Long

wServer.Close
'if we aren't waiting, exit sub
If Not WaitIncoming Then Exit Sub
created = False 'have we loaded a new socket?

If numSockets < progr + 1 Then 'if there is need to load a new socket, do it
    Load wSM(progr)
    Load wSD(progr)
    Load sDebug.lDebug(progr)
    'load also a label in the debug window
    sDebug.lDebug(progr).Top = progr * 24
    sDebug.lDebug(progr).Height = 25
    sDebug.lDebug(progr).Left = 0
    sDebug.lDebug(progr).Width = 289
    sDebug.lDebug(progr).Visible = True
    created = True
End If
'accept the incoming connection on command socket
wSM(progr).Accept (requestID)
'and listen on data socket
wSD(progr).Listen

If created Then numSockets = numSockets + 1

'increment ID
progr = progr + 1
'update label
lUsers.Caption = progr & "/" & nClients & " (waiting for connections...)"

'if we have reached the number of users...
If (progr = nClients) Then
    WaitIncoming = False
    lUsers.Caption = CStr(progr)
    'redim the pointer variables
    ReDim ptrs(nClients) As Long
    ReDim ptrl(nClients) As Long
    'ask the user for the file to load on all clients
    cdfs.Filter = "All files|*.*"
    cdfs.DialogTitle = "Open with UltraPad"
    cdfs.ShowOpen
    
    If Len(cdfs.FileName) <> 0 Then
        alltext = LoadFile(cdfs.FileTitle)
        'don't know if it's really necessary, but I wait for the connection
        'to be established on data socket
        For I = 0 To nClients - 1
            Do
                DoEvents
                Sleep 100
            Loop Until wSD(I).State = 7
            wSD(I).SendData "-1" & vbCrLf & alltext
            DoEvents
        Next I
    End If
    
    'activate all users with their ID
    Sleep 500
    For I = 0 To nClients - 1
        wSM(I).SendData "A" & I & "," & nClients & vbCrLf
    Next I
    tmrUpd.Enabled = True
    
    Exit Sub
End If

wServer.Listen
End Sub

Private Sub wServer_DataArrival(ByVal bytesTotal As Long)
datain = datain + bytesTotal
End Sub

Private Sub wServer_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If bytesRemaining = 0 Then dataout = dataout + bytesSent
End Sub

Private Sub wSM_Close(Index As Integer)
Dim I As Long

For I = 0 To nClients - 1
    wSD(I).Close
    wSM(I).Close
Next I
End Sub

Private Sub wSM_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim cmd As String, cmdp As String
Dim ncmd As Long, I As Long, idx As Long, temp As Long

datain = datain + bytesTotal
wSM(Index).GetData cmd, vbString

ncmd = getParams(cmd)

For I = 1 To ncmd
    cmdp = strParam(cmd, I)
    Select Case Left(cmdp, 1)
        'this sub is substantially equal to that in the client code.
        'it only does the checks
        Case "P"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        temp = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        'here is the check, if it's OK it notifies all the users,...
        If (checkPtr(idx, temp)) Then
            ptrs(idx) = temp
            ptrl(idx) = 0
            SendAll cmdp & vbCrLf
        Else
        'else it sends to the user the X command
            SendTo idx, "X" & vbCrLf
        End If
        DebugLabels

        Case "L"
        idx = CLng(Mid(Left(cmdp, InStr(cmdp, ",") - 1), 2))
        temp = CLng(Mid(cmdp, InStr(cmdp, ",") + 1))
        If (checkLen(idx, temp)) Then
            'check if it is too long, and eventually shorten it
            If ptrs(idx) + ptrl(idx) > Len(alltext) Then
                ptrl(idx) = Len(alltext) - ptrs(idx)
            Else
                ptrl(idx) = temp
            End If
            SendAll cmdp & vbCrLf
        Else
            SendTo idx, "X" & vbCrLf
        End If
        DebugLabels

        Case "C"
        'ping chat to all users
        SendAll cmdp & vbCrLf
    End Select
Next I

End Sub

Private Sub wSM_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If bytesRemaining = 0 Then dataout = dataout + bytesSent
End Sub

Private Sub DebugLabels()
Dim I As Integer

For I = 0 To nClients - 1
    sDebug.lDebug(I).Caption = "START " & ptrs(I) & ", LEN " & ptrl(I)
Next I
End Sub

