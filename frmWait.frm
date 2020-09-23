VERSION 5.00
Begin VB.Form cWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "attendere..."
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lText 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "cWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'close form if user hits ESC
If KeyCode = vbKeyEscape Then
    Me.Hide
    cMain.Enabled = True
End If
End Sub

