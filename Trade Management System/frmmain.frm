VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00AD8408&
   Caption         =   "Trade Management System"
   ClientHeight    =   11580
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   19320
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmmain.frx":0442
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      ForeColor       =   &H80000008&
      Height          =   11580
      Left            =   0
      ScaleHeight     =   11550
      ScaleWidth      =   3330
      TabIndex        =   0
      Top             =   0
      Width           =   3360
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   10200
         Width           =   2775
      End
      Begin VB.CommandButton cmdLogOut 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   9120
         Width           =   2775
      End
      Begin VB.CommandButton cmdPurChicks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Purchase Chicks"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CommandButton cmdCalculator 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7320
         Width           =   2775
      End
      Begin VB.CommandButton cmdSoldFish 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Sold Fish"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6240
         Width           =   2775
      End
      Begin VB.CommandButton cmdSoldChicken 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Sold Chicken"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   2775
      End
      Begin VB.CommandButton cmdSoldChicks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Sold Chicks"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   720
         Picture         =   "frmmain.frx":10716
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Managment System"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   2955
      End
   End
   Begin VB.Menu mnuChicks 
      Caption         =   "Chicks"
      Begin VB.Menu mnuPurchaseChicks 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSoldChicks 
         Caption         =   "Sold"
      End
      Begin VB.Menu mnuAliDedChicks 
         Caption         =   "Alived && Dead"
      End
   End
   Begin VB.Menu mnuPoultry 
      Caption         =   "Poultry"
      Begin VB.Menu mnuPurChicken 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSoldChicken 
         Caption         =   "Sold"
      End
      Begin VB.Menu mnuDailyIncChicken 
         Caption         =   "Daily Income"
      End
      Begin VB.Menu mnuDailyExpchicken 
         Caption         =   "Daily Expence"
      End
   End
   Begin VB.Menu mnuFish 
      Caption         =   "Fish"
      Begin VB.Menu mnuPurchaseFish 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSoldFish 
         Caption         =   "Sold"
      End
      Begin VB.Menu mnuDailyIncFish 
         Caption         =   "Daily Income"
      End
      Begin VB.Menu mnuDailyExpFish 
         Caption         =   "Daily Expence"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu mnuAdminSetting 
         Caption         =   "Setting"
      End
   End
   Begin VB.Menu mnuCalculator 
      Caption         =   "Calculator"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCalculator_Click()
    Dim Program As String, TaskID As Double
    Program = "calc.exe"
    On Error Resume Next
    AppActivate "Calculator"
    If Err <> 0 Then
    Err = 0
    TaskID = Shell(Program, 1)
    If Err <> 0 Then MsgBox "Can't start " & Program
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLogOut_Click()
    Unload Me
    LoginForm.Show
End Sub

Private Sub cmdPurChicks_Click()
    frm_BroughtChicks.Show
End Sub

Private Sub cmdSoldChicken_Click()
    frm_P_SoldChicken.Show
End Sub

Private Sub cmdSoldChicks_Click()
    frm_SoldChicks.Show
End Sub

Private Sub cmdSoldFish_Click()
    frm_F_SoldFish.Show
End Sub

Private Sub mnuAdminSetting_Click()
    frm_AdminSetting.Show
End Sub

Private Sub mnuAliDedChicks_Click()
    frm_AliveDead.Show
End Sub

Private Sub mnuBroughtChicks_Click()
    frm_BroughtChicks.Show
End Sub


Private Sub mnuCalculator_Click()
    Dim Program As String, TaskID As Double
    Program = "calc.exe"
    On Error Resume Next
    AppActivate "Calculator"
    If Err <> 0 Then
    Err = 0
    TaskID = Shell(Program, 1)
    If Err <> 0 Then MsgBox "Can't start " & Program
    End If
End Sub

Private Sub mnuDailyExpchicken_Click()
    frm_P_DailyExp.Show
End Sub

Private Sub mnuDailyExpFish_Click()
    frm_F_FishDailyExp.Show
End Sub

Private Sub mnuDailyIncChicken_Click()
    frm_P_DailyInco.Show
End Sub

Private Sub mnuDailyIncFish_Click()
    frm_F_FishDailyInco.Show
End Sub

Private Sub mnuPurchaseChicks_Click()
        frm_BroughtChicks.Show
End Sub

Private Sub mnuPurchaseFish_Click()
    frm_F_BroughtFish.Show
End Sub

Private Sub mnuPurChicken_Click()
frm_P_PurchaseChicken.Show
End Sub

Private Sub mnuSoldChicken_Click()
    frm_P_SoldChicken.Show
End Sub

Private Sub mnuSoldChicks_Click()
    frm_SoldChicks.Show
End Sub

Private Sub mnuSoldFish_Click()
    frm_F_SoldFish.Show
End Sub

Private Sub mnuTotalChicken_Click()
    frm_P_PoultryTotal.Show
End Sub

Private Sub mnuTotalChickenData_Click()
    frm_P_PoultryTotal.Show
End Sub

Private Sub mnuTotalChicks_Click()
    frm_TotalChicks.Show
End Sub

Private Sub mnuTotalChicksData_Click()
    frm_TotalChicks.Show
End Sub

Private Sub mnuTotalfish_Click()
    frm_F_FishTotal.Show
End Sub

Private Sub mnuTotalfishData_Click()
    frm_F_FishTotal.Show
End Sub
