VERSION 5.00
Begin VB.Form SchoolAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Assign"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDone 
      Caption         =   "Done"
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool1 
      Caption         =   "Basic School"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool2 
      Caption         =   "Average School"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool3 
      Caption         =   "Advanced School"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label LblNoSchool 
      Alignment       =   2  'Center
      Caption         =   "No schools are built!"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label LblUneducatedCount 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label LblUneducated 
      Alignment       =   1  'Right Justify
      Caption         =   "Uneducated:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "SchoolAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDone_Click()

SchoolAssign.Hide

End Sub

Private Sub CmdSchool1_Click()

If BasicSchoolProperties.LblStudentCount.Caption < 100 Then
    If LblUneducatedCount.Caption > 0 Then
        BasicSchoolProperties.LblStudentCount.Caption = BasicSchoolProperties.LblStudentCount.Caption + 1
        MainGame.LblUneducatedCount.Caption = MainGame.LblUneducatedCount.Caption - 1
        LblUneducatedCount.Caption = LblUneducatedCount.Caption - 1
            If MainGame.LblUneducatedCount.Caption = 0 Then
                MainGame.CmdAssignSchool.Visible = False
            End If
    End If
End If

If BasicSchoolProperties.LblStudentCount.Caption = 100 Then
MsgBox "The Basic School can only hold 100 students!"
End If

End Sub

Private Sub Form_Load()

LblUneducatedCount.Caption = MainGame.LblUneducatedCount.Caption

If Schools.CmdSchool1.Caption = "Basic School [Build]" Then
    CmdSchool1.Visible = False
End If
If Schools.CmdSchool2.Caption = "Average School [Build]" Then
    CmdSchool2.Visible = False
End If
If Schools.CmdSchool3.Caption = "Advanced School [Build]" Then
    CmdSchool3.Visible = False
End If

If Schools.CmdSchool1.Caption = "Basic School [Build]" Then
    If Schools.CmdSchool2.Caption = "Average School [Build]" Then
        If Schools.CmdSchool3.Caption = "Advanced School [Build]" Then
            LblNoSchool.Visible = True
        End If
    End If
End If

End Sub
