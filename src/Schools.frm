VERSION 5.00
Begin VB.Form Schools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schools"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDone 
      Caption         =   "Done"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool3 
      Caption         =   "Advanced School [Build]"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool2 
      Caption         =   "Average School [Build]"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdSchool1 
      Caption         =   "Basic School [Build]"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Schools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDone_Click()

Schools.Hide

End Sub

Private Sub CmdSchool1_Click()

If CmdSchool1.Caption = "Basic School" Then
    BasicSchoolProperties.Show
End If

If CmdSchool1.Caption = "Basic School [Build]" Then

    If AssignWorkers.LblTeacherUnempCount.Caption = 0 Then
        MsgBox "You need a teacher to teach!"
    End If

    If MainGame.LblWoodCount.Caption < 15 Then
        MsgBox "You need 15 wood to make a school!"
    End If

    If MainGame.LblWoodCount.Caption >= 15 Then
        If AssignWorkers.LblTeacherUnempCount.Caption > 0 Then
            CmdSchool1.Caption = "Basic School"
            MainGame.LblWoodCount.Caption = MainGame.LblWoodCount.Caption - 15
            AssignWorkers.LblTeacherUnempCount.Caption = AssignWorkers.LblTeacherUnempCount.Caption - 1
            AssignWorkers.LblTeacherEmpCount.Caption = AssignWorkers.LblTeacherEmpCount.Caption + 1
            BasicSchoolProperties.Show
            SchoolAssign.CmdSchool1.Visible = True
            SchoolAssign.LblNoSchool.Visible = False
        End If
    End If
End If

End Sub
