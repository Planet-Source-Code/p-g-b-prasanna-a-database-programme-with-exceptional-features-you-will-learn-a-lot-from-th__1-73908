VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login"
   ClientHeight    =   1875
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4065
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1107.812
   ScaleMode       =   0  'User
   ScaleWidth      =   3816.815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   68
      Width           =   3855
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   40
         ScaleHeight     =   1455
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   120
         Width           =   3735
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            IMEMode         =   3  'DISABLE
            Left            =   1320
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   7
            Top             =   520
            Width           =   2325
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2500
            MouseIcon       =   "frmLogin.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   1020
            Width           =   1120
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1300
            MouseIcon       =   "frmLogin.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   1020
            Width           =   1140
         End
         Begin VB.ComboBox cmbSelectUser 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   2340
         End
         Begin VB.Image Image2 
            Height          =   510
            Left            =   120
            Picture         =   "frmLogin.frx":02B0
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   540
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&User Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   150
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Dim rstusernames As ADODB.Recordset
Dim password As String

Private Sub cmdCancel_Click()
End
End Sub
Private Sub cmdOk_Click()
Check_Password
End Sub
Private Sub Form_Activate()
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

On Error GoTo db_Error
Set rstusernames = New ADODB.Recordset
rstusernames.CursorLocation = adUseClient
rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly

Add_User_Names_to_Combo
Retrieve_Last_User

On Error GoTo Err
If ret_user <> "" Then
    cmbSelectUser.Text = ret_user
Else
    cmbSelectUser.Text = cmbSelectUser.List(0)
End If
Exit Sub

Err:
If cmbSelectUser.ListCount > 0 Then
    cmbSelectUser.Text = cmbSelectUser.List(0)
    Exit Sub
End If

db_Error:
MsgBox "Database Error: " & Err.Number & "." & vbCrLf & "System is unable to continue." & vbCrLf & "Replace the database with a backup.", vbCritical
Reg_Obj.RegDelete (Database_Path_Store)
End
End Sub
Public Sub Add_User_Names_to_Combo()
If rstusernames.RecordCount > 0 Then
    Do While Not rstusernames.EOF
        cmbSelectUser.AddItem rstusernames("USER_NAME")
        rstusernames.MoveNext
    Loop
End If
End Sub
Public Sub Check_Password()
 Set rstgetpassword = New ADODB.Recordset
 rstgetpassword.CursorLocation = adUseClient
 rstgetpassword.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbSelectUser & "'", dbcon, adOpenStatic, adLockOptimistic
 password = Decrypt(rstgetpassword("PASSWORD"))
 intaccount_type = rstgetpassword("TYPE")
     
    If Not IsNull(rstgetpassword("READ_P")) Then
         user_read_privilege = rstgetpassword("READ_P")
    Else
         user_read_privilege = 0
    End If
    
    If Not IsNull(rstgetpassword("WRITE_P")) Then
         user_write_privilege = rstgetpassword("WRITE_P")
    Else
         user_write_privilege = 0
    End If
    
    User = cmbSelectUser
    
    If txtPassword = password Then
        Store_User
        frmSplash.Show
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", vbCritical, "Login"
        txtPassword.SetFocus
        With txtPassword
            .SetFocus
            .SelStart = 0
            .SelLength = Len(txtPassword)
        End With
    End If
'-------------------------------------------------------------------
End Sub

Public Sub Store_User()
On Error Resume Next
Call Write_Registry(Loggeduser, User)
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Unload frmStyle
rstusernames.Close
Set rstusernames = Nothing
Set Me.Picture = Nothing
End Sub

