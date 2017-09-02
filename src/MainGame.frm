VERSION 5.00
Begin VB.Form MainGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Game"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9825
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmChecks 
      Caption         =   "Checks"
      Height          =   2055
      Left            =   9960
      TabIndex        =   29
      Top             =   3480
      Width           =   2295
      Begin VB.CheckBox CheckFirstYearAtSchool 
         Caption         =   "FirstYearAtSchool"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FrmDMY 
      Caption         =   "D/M/Y"
      Height          =   615
      Left            =   4440
      TabIndex        =   22
      Top             =   120
      Width           =   3735
      Begin VB.Label LblMonthName 
         Caption         =   "June"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblYear 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "/"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblMonth 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "/"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblDay 
         Alignment       =   2  'Center
         Caption         =   "25"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer TimerDateTimeYear 
      Interval        =   1000
      Left            =   10800
      Top             =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8520
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame FrmSupplies 
      Caption         =   "Supplies"
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton CmdSuppliesClose 
         Caption         =   "x"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
      Begin VB.Label LblWoodCount 
         Caption         =   "25"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblWood 
         Caption         =   "Wood:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdSupplies 
      Caption         =   "Supplies"
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FrmTownInformation 
      Caption         =   "Town Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton CmdSchools 
         Caption         =   "Schools"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Houses"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdTownInfoClose 
         Caption         =   "x"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CommandButton CmdTownInfo 
      Caption         =   "Town Information"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1815
   End
   Begin VB.Timer TimerUneducated 
      Interval        =   1000
      Left            =   10800
      Top             =   1440
   End
   Begin VB.CommandButton CmdInfo 
      Caption         =   "General Information"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FrmInfoTab 
      Caption         =   "General Information Tab"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
      Begin VB.CommandButton CmdAssignSchool 
         Caption         =   "Assign School"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdAssignWork 
         Caption         =   "Assign Work"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdInfoClose 
         Caption         =   "x"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.Label LblUneducatedCount 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.Label LblUneducated 
         Caption         =   "Uneducated:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblWorkers 
         Caption         =   "Unemployed:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LblWorkersCount 
         Caption         =   "2"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer TimerConstant 
      Interval        =   1000
      Left            =   10800
      Top             =   840
   End
   Begin VB.Label LblTimerUneducated 
      Caption         =   "0"
      Height          =   255
      Left            =   11400
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label LblTimerConstant 
      Caption         =   "0"
      Height          =   255
      Left            =   11400
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "MainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SC_CLOSE As Long = &HF060&
Private Const MF_BYCOMMAND = &H0&
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Sub CmdAssignSchool_Click()

If LblUneducatedCount.Caption > 0 Then
    SchoolAssign.Show
End If

End Sub

Private Sub CmdInfo_Click()

FrmInfoTab.Visible = True
CmdInfo.Visible = False
FrmTownInformation.Visible = False
CmdTownInfo.Visible = True
CmdSupplies.Visible = True
FrmSupplies.Visible = False

End Sub

Private Sub CmdInfoClose_Click()

CmdInfo.Visible = True
FrmInfoTab.Visible = False

End Sub

Private Sub CmdAssignWork_Click()

If LblWorkersCount.Caption > 0 Then
    AssignWorkers.Show
End If

End Sub

Private Sub CmdSchools_Click()

Schools.Show

End Sub

Private Sub CmdSupplies_Click()

FrmSupplies.Visible = True
CmdTownInfo.Visible = True
CmdInfo.Visible = True
CmdSupplies.Visible = False
FrmInfoTab.Visible = False
FrmTownInformation.Visible = False

End Sub

Private Sub CmdSuppliesClose_Click()

CmdSupplies.Visible = True
FrmSupplies.Visible = False

End Sub

Private Sub CmdTownInfo_Click()

FrmInfoTab.Visible = False
CmdTownInfo.Visible = False
FrmTownInformation.Visible = True
CmdInfo.Visible = True
CmdSupplies.Visible = True
FrmSupplies.Visible = False

End Sub

Private Sub CmdTownInfoClose_Click()

CmdTownInfo.Visible = True
FrmTownInformation.Visible = False

End Sub

Private Sub Command1_Click()

Houses.Show

End Sub


Private Sub Command2_Click()

ExitGame.Show
Me.Hide

End Sub

Private Sub Form_Load()

DeleteMenu GetSystemMenu(Me.hWnd, False), SC_CLOSE, MF_BYCOMMAND

FrmInfoTab.Visible = False

End Sub



Private Sub LblWood_Click()

MsgBox "Gathering wood is simple. Both unemployed and uneducated people can gather wood."

End Sub

Private Sub TimerConstant_Timer()

LblTimerConstant.Caption = LblTimerConstant.Caption + 1

If LblTimerConstant.Caption = 2 Then
    LblTimerConstant.Caption = 0
End If

End Sub

Private Sub TimerDateTimeYear_Timer()

'keep the days rolling
LblDay.Caption = LblDay.Caption + 1

'/////////////////////////////////////
'/////////Months of the year//////////
'/////////////////////////////////////

'Switch to January
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 12 Then
        LblDay.Caption = 1
        LblMonth.Caption = 1
        LblMonthName.Caption = "January"
        LblYear.Caption = LblYear.Caption + 1
        
    End If
End If

'Switch to February
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 1 Then
        LblDay.Caption = 1
        LblMonth.Caption = 2
        LblMonthName.Caption = "February"
    End If
End If

'Switch to March
If LblDay.Caption = 29 Then
    If LblMonth.Caption = 2 Then
        LblDay.Caption = 1
        LblMonth.Caption = 3
        LblMonthName.Caption = "March"
    End If
End If

'Switch to April
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 3 Then
        LblDay.Caption = 1
        LblMonth.Caption = 4
        LblMonthName.Caption = "April"
    End If
End If

'Switch to May
If LblDay.Caption = 31 Then
    If LblMonth.Caption = 4 Then
        LblDay.Caption = 1
        LblMonth.Caption = 5
        LblMonthName.Caption = "May"
    End If
End If

'Switch to June
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 5 Then
        LblDay.Caption = 1
        LblMonth.Caption = 6
        LblMonthName.Caption = "June"
    End If
End If

'Switch to July
If LblDay.Caption = 31 Then
    If LblMonth.Caption = 6 Then
        LblDay.Caption = 1
        LblMonth.Caption = 7
        LblMonthName.Caption = "July"
    End If
End If

'Switch to August
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 7 Then
        LblDay.Caption = 1
        LblMonth.Caption = 8
        LblMonthName.Caption = "August"
    End If
End If

'Switch to September
If LblDay.Caption = 31 Then
    If LblMonth.Caption = 8 Then
        LblDay.Caption = 1
        LblMonth.Caption = 9
        LblMonthName.Caption = "September"
    End If
End If

'Switch to October
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 9 Then
        LblDay.Caption = 1
        LblMonth.Caption = 10
        LblMonthName.Caption = "October"
    End If
End If

'Switch to November
If LblDay.Caption = 31 Then
    If LblMonth.Caption = 10 Then
        LblDay.Caption = 1
        LblMonth.Caption = 11
        LblMonthName.Caption = "November"
    End If
End If

'Switch to December
If LblDay.Caption = 32 Then
    If LblMonth.Caption = 11 Then
        LblDay.Caption = 1
        LblMonth.Caption = 12
        LblMonthName.Caption = "December"
    End If
End If

End Sub

Private Sub TimerUneducated_Timer()


LblTimerUneducated.Caption = LblTimerUneducated + 1

If LblTimerUneducated.Caption = 11 Then
    LblTimerUneducated.Caption = 0
    LblUneducatedCount.Caption = LblUneducatedCount.Caption + 1
    SchoolAssign.LblUneducatedCount.Caption = LblUneducatedCount.Caption
End If

If LblUneducatedCount.Caption > 0 Then
    Call SchoolTime
End If

End Sub

Private Sub SchoolTime()

If LblMonthName.Caption = "September" Then
    CmdAssignSchool.Visible = False
    SchoolAssign.Hide
End If

If CheckFirstYearAtSchool.Value = 0 Then
    MsgBox "It's now June! You can sign up unemployed people for school during the months June, July, and August provided that you have uneducated people ;)"
    CheckFirstYearAtSchool.Value = 1
End If

If LblMonthName.Caption = "June" Then
    If LblUneducatedCount.Caption > 0 Then
        CmdAssignSchool.Visible = True
    End If
End If

If LblMonthName.Caption = "July" Then
    If LblUneducatedCount.Caption > 0 Then
        CmdAssignSchool.Visible = True
    End If
End If

If LblMonthName.Caption = "August" Then
    If LblUneducatedCount.Caption > 0 Then
        CmdAssignSchool.Visible = True
    End If
End If

End Sub
