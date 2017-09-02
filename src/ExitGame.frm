VERSION 5.00
Begin VB.Form ExitGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "   Cancel     (back to game)"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Don't save (quit)"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmdSaveGame 
      Caption         =   "Save Game"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to quit without saving?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "ExitGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
MainGame.Show
ExitGame.Hide
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdSaveGame_Click()
SaveGame.Show
Me.Hide
End Sub
