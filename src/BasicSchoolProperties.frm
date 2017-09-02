VERSION 5.00
Begin VB.Form BasicSchoolProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic School Properties"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdUpgrade 
      Caption         =   "Upgrade"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Upgrade 4"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Upgrade 3"
      Height          =   195
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Upgrade 2"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Upgrade 1"
      Height          =   195
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "Done"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label LblLevel 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "/   100"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LblStudentCount 
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LblStudents 
      Caption         =   "Students:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "BasicSchoolProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDone_Click()

BasicSchoolProperties.Hide

End Sub

Private Sub CmdUpgrade_Click()

If Chk1.Value = 0 Then
    Chk1.Value = 1
    LblLevel.Caption = "Level 1"
    Exit Sub
End If

If Chk2.Value = 0 Then
    Chk2.Value = 1
    LblLevel.Caption = "Level 2"
    Exit Sub
End If

If Chk3.Value = 0 Then
    Chk3.Value = 1
    LblLevel.Caption = "Level 3"
    Exit Sub
End If

If Chk4.Value = 0 Then
    Chk4.Value = 1
    LblLevel.Caption = "Level 4"
    Exit Sub
End If

If Chk4.Value = 1 Then
    MsgBox "The school has reached its maximum upgrade!"
End If

End Sub
