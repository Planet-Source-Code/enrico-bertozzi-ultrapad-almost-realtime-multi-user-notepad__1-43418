VERSION 5.00
Begin VB.Form sDebug 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lDebug 
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "sDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'another simple form to display some debug info
Private Sub Form_Load()

End Sub
