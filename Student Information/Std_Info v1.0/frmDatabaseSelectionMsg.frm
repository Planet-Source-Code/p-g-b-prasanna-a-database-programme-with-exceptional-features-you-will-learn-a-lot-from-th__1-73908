VERSION 5.00
Begin VB.Form frmDatabaseSelectionMsg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Configuration..."
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "frmDatabaseSelectionMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4420
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   6255
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4155
         Left            =   50
         ScaleHeight     =   4155
         ScaleWidth      =   6135
         TabIndex        =   1
         Top             =   120
         Width           =   6135
         Begin VB.OptionButton optShared 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Shared Location:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   1680
            Width           =   3855
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtLocal 
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   840
            Width           =   4695
         End
         Begin VB.OptionButton optLocalMapped 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Local Computer or Network Drive:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtShared 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   10
            Top             =   2280
            Width           =   4575
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "C&ancel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CommandButton cmdConfigure 
            Caption         =   "&Configure..."
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please locate the database."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   120
            Width           =   2340
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1200
            TabIndex        =   7
            Top             =   2955
            Width           =   465
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For any technical issue contact me on:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1200
            TabIndex        =   6
            Top             =   2760
            Width           =   2760
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pgbsoft@gmail.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   2955
            Width           =   1410
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxx - xxxxxxx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3975
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   2760
            Width           =   1050
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "(Example: \\server\dbshared folder\db_stdinfo.mdb)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   3
            Top             =   2040
            Width           =   3855
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   120
            Picture         =   "frmDatabaseSelectionMsg.frx":06A6
            Top             =   120
            Width           =   750
         End
         Begin VB.Line Line1 
            X1              =   1200
            X2              =   5880
            Y1              =   3480
            Y2              =   3480
         End
      End
   End
End
Attribute VB_Name = "frmDatabaseSelectionMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Private Sub cmdBrowse_Click()
intbrowseoption = 2
frmBF.Show 1
End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub cmdConfigure_Click()
On Error Resume Next
If intbrowseoption = 1 Then
    If UCase(Right(txtShared, 4)) <> UCase(".mdb") Then
       MsgBox "Invalid Database.", vbExclamation
       txtShared.SetFocus
       SendKeys "{HOME}+{END}"
       Exit Sub
    End If
End If

If intbrowseoption = 1 Then
   Located_Database = txtShared
   Call Write_Registry(Database_Path_Store, Located_Database)
   Database_Path = Located_Database
ElseIf intbrowseoption = 2 Then
   Located_Database = txtLocal
   Call Write_Registry(Database_Path_Store, Located_Database)
   Database_Path = Located_Database
End If
   intdbproceed = 1
   Unload Me
   openDatabase
End Sub

Private Sub cmdConfigure_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Form_Activate()
On Error Resume Next
If intbrowseoption = 1 Then
    optShared.Value = True
ElseIf intbrowseoption = 2 Then
    optLocalMapped.Value = True
ElseIf intbrowseoption = 0 Then
    optLocalMapped.Value = True
End If
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture
cmdConfigure.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Me.Picture = Nothing
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub lblEmail_Click()
Mail_Me "pgbsoft@gmail.com"
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = True
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub optLocalMapped_Click()
intbrowseoption = 2
txtLocal.Enabled = True
txtShared = ""
txtShared.Enabled = False
cmdBrowse.Enabled = True
cmdBrowse.SetFocus
End Sub

Private Sub optShared_Click()
'cmdSelectLocal.Enabled = False
'cmdSelectMapped.Enabled = False
'cmdConvert.Enabled = False
intbrowseoption = 1
txtLocal.Enabled = False
cmdBrowse.Enabled = False
txtShared.Enabled = True
txtLocal = ""
txtShare = ""
On Error Resume Next
txtShared.SetFocus
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub txtLocal_Change()
If txtLocal <> "" Then
    cmdConfigure.Enabled = True
Else
    cmdConfigure.Enabled = False
End If
End Sub

Private Sub txtShared_Change()
If txtShared <> "" Then
    cmdConfigure.Enabled = True
Else
    cmdConfigure.Enabled = False
End If
End Sub

Private Sub txtShared_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub
