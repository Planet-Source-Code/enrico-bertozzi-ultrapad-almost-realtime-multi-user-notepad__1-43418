Attribute VB_Name = "QSync"
Option Explicit

'the UltraPad server's main job is to manage connections and ping all commands
'and data to all clients. It also assigns areas to clients and checks collisions.
'this server implementation doesn't support user log-in / log-out. that signifies
'if a user exits, all the session is closed, and when a session starts, all the
'user have to be connected

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global nClients As Long 'number of clients
Global WaitIncoming As Boolean 'flag that indicates if the server is awaiting connections
                            'from clients
Global FileSaved As Boolean 'signals if the file has been saved or not
Global numSockets As Long 'number of sockets currently loaded
Global progr As Long 'progressive ID assignment

Global datain As Long, dataout As Long 'data in/out registers (statistics)

Global ptrs() As Long 'pointers
Global ptrl() As Long

Global alltext As String 'contains the file

Public Const Version As String = "UltraPad Server 3.0"

Public Function strParam(str As String, ByVal n As Long) As String
'see Client for details
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
Dim I%, p%

p = 0

For I = 1 To Len(str)
    If Mid(str, I, 2) = vbCrLf Then p = p + 1
Next I

getParams = p
    
End Function

Function LoadFile(file As String) As String
'load a text file to a string
Open file For Input As #1
LoadFile = Input(LOF(1), 1)
Close #1
End Function

Sub SaveFile(file As String, data As String)
'saves a string to a file
Open file For Output As #2
Print #2, data;
Close #2
End Sub

Public Sub SendExcept(ByVal except As Long, text As String)
'sends a string to all users except one. In fact it is not used, but
'I leave this anyway

Dim I As Long

For I = 0 To nClients - 1
    If I <> except Then sMain.wSM(I).SendData text
    DoEvents
Next I
End Sub

Public Sub SendAll(text As String)
'sends to all users a string

Dim I As Long

For I = 0 To nClients - 1
    sMain.wSM(I).SendData text
    DoEvents
Next I
End Sub

Public Sub SendTo(ByVal dest As Long, text As String)
'sends to an user a string

sMain.wSM(dest).SendData text
DoEvents
End Sub

Public Function checkPtr(n As Long, ptr As Long) As Boolean
'this checks if the specified pointer collides with ones from another clients

Dim I As Long

checkPtr = True
For I = 0 To nClients - 1
    If I <> n Then
        If (ptr > ptrs(I) And ptr < ptrs(I) + ptrl(I)) Then checkPtr = False
    End If
Next I
End Function

Public Function checkLen(n As Long, leng As Long) As Boolean
'this checks the area if it collides with another user's one

Dim I As Long

checkLen = True
For I = 0 To nClients - 1
    If I <> n Then
        If (leng + ptrs(n) > ptrs(I) And leng + ptrs(n) < ptrs(I) + ptrl(I)) Then checkLen = False
    End If
Next I
End Function
