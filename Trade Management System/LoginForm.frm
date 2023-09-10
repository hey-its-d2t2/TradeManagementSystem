VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00AA9D23&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   4695
   ClientLeft      =   7275
   ClientTop       =   3345
   ClientWidth     =   7230
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00AD8408&
      Height          =   570
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "UserName"
      DataSource      =   "Adodc1"
      Height          =   570
      Left            =   4320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"LoginForm.frx":0442
      OLEDBString     =   $"LoginForm.frx":04FC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from UserDetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&Clear "
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
      Left            =   600
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clear Fields "
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00AD8408&
      Caption         =   "&Ok"
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
      Left            =   5040
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ok"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00000000&
      Picture         =   "LoginForm.frx":05B6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA9D23&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "#"
      TabIndex        =   0
      ToolTipText     =   "Enter Password"
      Top             =   2325
      Width           =   4095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      X1              =   2880
      X2              =   6960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblMsgP 
      AutoSize        =   -1  'True
      BackColor       =   &H00AA9D23&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0091DCF5&
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label lblUID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0091DCF5&
      BackStyle       =   0  'Transparent
      Caption         =   "User Id : "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lblLogin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Login "
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   1200
   End
   Begin VB.Line Line3 
      X1              =   7200
      X2              =   7200
      Y1              =   840
      Y2              =   4680
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7200
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   4680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
        Dim Reply As Integer
        Reply = MsgBox("Do you want ot Exit.", vbOKCancel + vbInformation, "Exit ?")
        If Reply = vbOK Then
            End
        End If
End Sub

Private Sub cmdClear_Click()
  ' txtUId.Text = ""
    txtPassword.Text = ""
    lblMsgU.Caption = ""
    lblMsgP.Caption = ""
End Sub

Private Sub cmdExit_Click()
       Dim Reply As Integer
        Reply = MsgBox("Do you want ot Exit.", vbYesNo + vbInformation, "Exit ?")
        If Reply = vbYes Then
            End
        End If
End Sub


Private Sub cmdOk_Click()
   'Code for Responsive
   If txtPassword.Text = "" Then
             lblMsgP.Caption = "Password can't blank."
         End If
    'Code for User Validation
    If txtPassword.Text <> "" Then
            Adodc1.RecordSource = "select * from UserDetails where UserName='" & Combo1.Text & "' and Password='" & txtPassword.Text & "'"
            Adodc1.Refresh
            If Adodc1.Recordset.RecordCount = 0 Then
                   lblMsgP.Caption = "Invalid Password."
            Else
                Unload Me
                frmmain.Show
        End If
    End If
    Adodc1.Recordset.Close
End Sub


Private Sub Form_Load()
    Combo1.Clear
    Adodc1.RecordSource = "select * from UserDetails"
    Adodc1.Refresh
    While Not Adodc1.Recordset.EOF
        Combo1.AddItem Adodc1.Recordset.Fields(1).Value
        Adodc1.Recordset.MoveNext
    Wend
    Combo1.Text = Combo1.List(0)
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
     lblMsgP.Caption = ""
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtPassword.Text = "" Then
            lblMsgP.Caption = "Password can't blank."
        End If
    cmdOk.SetFocus
    End If
End Sub





