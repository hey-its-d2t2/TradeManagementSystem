VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_BroughtChicks 
   Appearance      =   0  'Flat
   BackColor       =   &H00AA9D23&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "         Purchase Chicks"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   15135
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   25
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      CalendarBackColor=   11181347
      CalendarTitleBackColor=   11371528
      Format          =   71303169
      CurrentDate     =   44792
   End
   Begin VB.TextBox Text1 
      DataField       =   "SupName"
      DataSource      =   "Adodc1"
      Height          =   510
      Left            =   12240
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Bindings        =   "frm_BroughtChicks.frx":0000
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   5445
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   8493
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16776682
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Purchase Chicks Details"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Next Recoard"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete Recoard"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Previous Record"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add New Data"
      Top             =   2400
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8520
      Top             =   4800
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DSNTMS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DSNTMS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BroughtChicks"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Update fields"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtPice 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1875
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   360
      TabIndex        =   17
      Top             =   0
      Width           =   6615
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00AD8408&
         Height          =   510
         ItemData        =   "frm_BroughtChicks.frx":0015
         Left            =   2520
         List            =   "frm_BroughtChicks.frx":0025
         TabIndex        =   5
         Text            =   "Cash"
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5640
         Y1              =   3880
         Y2              =   3880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5640
         Y1              =   3160
         Y2              =   3160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5640
         Y1              =   1720
         Y2              =   1720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pice :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   22
         Top             =   1290
         Width           =   1380
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (Per Pices)  :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   21
         Top             =   1995
         Width           =   2160
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   20
         Top             =   2685
         Width           =   1200
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Avg. Weight :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   19
         Top             =   3390
         Width           =   1740
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Mode :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save Recoard"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   7680
      TabIndex        =   14
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtSupMob 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtSupName 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5880
         Y1              =   1720
         Y2              =   1720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         X1              =   2520
         X2              =   5880
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   2025
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AA9D23&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Mob.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frm_BroughtChicks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
     Dim Reply As Integer
        Reply = MsgBox("Do you want delete.", vbYesNo + vbInformation, "Exit ?")
        If Reply = vbYes Then
            Adodc1.Recordset.Delete
             MsgBox "record deleted.", vbOKOnly + vbInformation, "Information"
        End If
End Sub

Private Sub cmdNew_Click()
    ClearAll
End Sub

Private Sub getData()

    If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.BOF Then
               MsgBox "You reached at first data", vbOKOnly + vbInformation, "Information"
        Else
            MsgBox "You reached at last data", vbOKOnly + vbInformation, "Information"
        End If
    Else
    With Adodc1.Recordset
        DTPicker1.Value = .Fields(0).Value
        txtSupName.Text = .Fields(1).Value
        txtSupMob.Text = .Fields(2).Value
        txtPice.Text = .Fields(3).Value
        txtRate.Text = .Fields(4).Value
        txtAmount.Text = .Fields(5).Value
        txtWeight.Text = .Fields(6).Value
        Combo1.Text = .Fields(7).Value
    End With
    End If
End Sub

Private Sub cmdNext_Click()
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveLast
    getData
Else
    Adodc1.Recordset.MoveNext
    getData
End If

End Sub

Private Sub cmdPrev_Click()
    If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
        getData
    Else
        Adodc1.Recordset.MovePrevious
         getData
    End If
    
    
End Sub

Private Sub cmdSave_Click()
Adodc1.Recordset.AddNew
With Adodc1.Recordset
    .Fields(0).Value = DTPicker1.Value
    .Fields(1).Value = txtSupName.Text
    .Fields(2).Value = txtSupMob.Text
    .Fields(3).Value = txtPice.Text
    .Fields(4).Value = txtRate.Text
    .Fields(5).Value = txtAmount.Text
    .Fields(6).Value = txtWeight.Text
    .Fields(7).Value = Combo1.Text
End With
Adodc1.Recordset.Update
Adodc1.Refresh
 MsgBox "Record saved.", vbOKOnly + vbInformation, "Information"
End Sub

Private Sub cmdUpdate_Click()
     With Adodc1.Recordset
    .Fields(0).Value = DTPicker1.Value
    .Fields(1).Value = txtSupName.Text
    .Fields(2).Value = txtSupMob.Text
    .Fields(3).Value = txtPice.Text
    .Fields(4).Value = txtRate.Text
    .Fields(5).Value = txtAmount.Text
    .Fields(6).Value = txtWeight.Text
    .Fields(7).Value = Combo1.Text
End With
Adodc1.Recordset.Update
Adodc1.Refresh
 MsgBox "Details updated.", vbOKOnly + vbInformation, "Information"
End Sub

Private Sub Form_Load()
     ClearAll
End Sub
Private Sub ClearAll()
    txtPice.Text = ""
    txtRate.Text = ""
    txtAmount.Text = ""
    txtWeight.Text = ""
    txtSupName.Text = ""
    txtSupMob.Text = ""
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
      txtWeight.SetFocus
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    txtPice.SetFocus
    End If
End Sub

Private Sub txtPice_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        txtRate.SetFocus
    End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    txtAmount.Text = Val(txtPice.Text()) * Val(txtRate.Text())
    If KeyAscii = vbKeyReturn Then
         txtWeight.SetFocus
     End If
End Sub



Private Sub txtSupName_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
         txtSupMob.SetFocus
     End If
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Combo1.SetFocus
    End If
End Sub
