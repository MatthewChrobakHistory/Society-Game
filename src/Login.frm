VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmLoad 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   7680
      TabIndex        =   11
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdExit 
         Caption         =   "x"
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton CmdLoad1 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad2 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad3 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton CmdLoad4 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.Frame FramePlay 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton CmdExit2 
         Caption         =   "x"
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load a Saved Game"
         Height          =   615
         Left            =   600
         TabIndex        =   9
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton CmdPlayGame 
         Caption         =   "Play New Game"
         Height          =   615
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Or"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "Enter"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.OLE OLE1 
      BorderStyle     =   0  'None
      Class           =   "Paint.Picture"
      Height          =   1575
      Left            =   960
      OleObjectBlob   =   "Login.frx":0000
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Welcome to my game!"
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Height          =   3375
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label LblBackround2 
      BackColor       =   &H80000010&
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblBackround 
      BackColor       =   &H80000011&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEnter_Click()

FramePlay.Left = lblBackround.Left

End Sub

Private Sub CmdExit_Click()
FrmLoad.Left = 7680
End Sub

Private Sub CmdExit2_Click()
End
End Sub

Private Sub CmdPlayGame_Click()
Login.Hide
MainGame.Show
End Sub

Private Sub Command2_Click()
FrmLoad.Left = FramePlay.Left
End Sub
