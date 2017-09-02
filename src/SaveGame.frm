VERSION 5.00
Begin VB.Form SaveGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmLoad 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdLoad4 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad3 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad2 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad1 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "x"
         Height          =   195
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "SaveGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
ExitGame.Show
SaveGame.Hide
End Sub

Private Sub CmdLoad1_Click()

If Dir(App.Path & "\Game Data\Saves\Save1\") = "" Then
    MkDir (App.Path & "Game Data\Saves\Save1\")
End If
End Sub
