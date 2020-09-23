Attribute VB_Name = "Other"
Option Explicit

'UltraPad works using two sockets per user. The command socket, where command
'are issued and separated by a vbCrLf, and the data socket, where data is simply
'sent with USERID & vbCrLf for the header (needed to know where is the data from).
'The editor is on a little different basis: one cannot edit what he wants, but
'all clients must share the space, and once a client clicks in the main textbox,
'that area is assigned to the client. If the area is already occupied, the programs
'emits a beep.
'These "area" data is stored on two arrays: PTRS is the starting character of area of
'each user. PTRL is the lenght.
'To communicate, each client is assigned an ID (myId), so other clients know what
'pointer to modify if they receive some command.
'In UltraPad the server has only to eventually save/load the file and retransmit
'the messages. I did it separately only to tidy up all.... :)

Global nClients As Long 'total number of clients connected to the server
Global myId As Long 'user ID

Global w As Long, h As Long 'window width and height, used for layout

Global ptrs() As Long 'starting area pointer
Global ptrl() As Long 'area length

Global editchanged As Boolean 'flag that stores if the edit txtbox has modified
Global mainenabled As Boolean 'flag indicating if the client is connected

Public Const Version As String = "UltraPad Executive 3.0"

Public word(0 To 31) As String 'syntax highlight

Public Function strParam(str As String, ByVal n As Long) As String
'Used to get a particular parameter in a string, separated with vbCrLfs
'This function and the one below are here only because Winsock. If I send
'two strings on a socket, even using two separate commands, the client may
'receive them together and glued

Dim I%, sp%, ss%

sp = 1

For I = 1 To Len(str)
    If sp = n Then Exit For
    If Mid(str, I, 2) = vbCrLf Then sp = sp + 1
Next I

ss = I

For I = ss To Len(str)
    If Mid(str, I, 2) = vbCrLf Then Exit For
Next I

strParam = Mid(str, IIf(n = 1, ss, ss + 1), IIf(n = 1, I - ss, I - ss - 1))
End Function

Public Function getParams(str As String) As Integer
'To get the total number of parameters of a string

Dim I%, p%

p = 0

For I = 1 To Len(str)
    If Mid(str, I, 2) = vbCrLf Then p = p + 1
Next I

getParams = p
    
End Function

Sub ShowWait(text As String)
'Show the "wait form".

cWait.lText.Caption = text
cWait.Show 0, cMain
cMain.Enabled = False
DoEvents
End Sub

Sub WaitText(text As String)
cWait.lText.Caption = text
End Sub

Sub HideWait()
cWait.Hide
cMain.Enabled = True
End Sub

Public Sub Send(text As String)
'shortly, send text on the socket

cMain.wClient.SendData text
DoEvents
End Sub
