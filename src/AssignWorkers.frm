VERSION 5.00
Begin VB.Form AssignWorkers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Workers"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmWoodcutter 
      Caption         =   "Woodcutters"
      Height          =   975
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton CmdEmployWoodcutter 
         Caption         =   "Employ"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblWoodcutterEmpCount 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblWoodcutterEmp 
         Caption         =   "Emp:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblWoodcutterUnemp 
         Caption         =   "Unemp:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblWoodcutterUnempCount 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "Done"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Find"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox CmboFindAProfession 
      Height          =   315
      ItemData        =   "AssignWorkers.frx":0000
      Left            =   120
      List            =   "AssignWorkers.frx":000A
      TabIndex        =   7
      Text            =   "Find a profession"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame FrmTeacher 
      Caption         =   "Teachers"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton CmdEmployTeacher 
         Caption         =   "Employ"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblTeacherUnempCount 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label LblTeacherEmpCount 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblTeacherUnemp 
         Caption         =   "Unemp:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblTeacherEmp 
         Caption         =   "Emp:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label LblUnemployedCount 
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblUnemployed 
      Caption         =   "Unemployed:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "AssignWorkers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDone_Click()

AssignWorkers.Hide

End Sub

Private Sub CmdEmployTeacher_Click()

If LblUnemployedCount.Caption > 0 Then
    LblTeacherUnempCount.Caption = LblTeacherUnempCount.Caption + 1
    LblUnemployedCount.Caption = LblUnemployedCount.Caption - 1
    MainGame.LblWorkersCount.Caption = MainGame.LblWorkersCount.Caption - 1
End If

If MainGame.LblWorkersCount.Caption = 0 Then
    MainGame.CmdAssignWork.Visible = False
End If

End Sub

Private Sub CmdEmployWoodcutter_Click()

If LblUnemployedCount.Caption > 0 Then
    LblWoodcutterUnempCount.Caption = LblWoodcutterUnempCount.Caption + 1
    LblUnemployedCount.Caption = LblUnemployedCount.Caption - 1
    MainGame.LblWorkersCount.Caption = MainGame.LblWorkersCount.Caption - 1
End If

If MainGame.LblWorkersCount.Caption = 0 Then
    MainGame.CmdAssignWork.Visible = False
End If

End Sub

Private Sub CmdGo_Click()

'Hide all the tabs
FrmTeacher.Visible = False
FrmWoodcutter.Visible = False

'Teacher
If CmboFindAProfession.Text = "Teacher" Then
    FrmTeacher.Visible = True
End If

'Woodcutter
If CmboFindAProfession.Text = "Woodcutter" Then
    FrmWoodcutter.Visible = True
End If

End Sub

Private Sub Form_Load()

LblUnemployedCount.Caption = MainGame.LblWorkersCount

End Sub
