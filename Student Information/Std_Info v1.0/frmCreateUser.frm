VERSION 5.00
Begin VB.Form frmCreateUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create User..."
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmCreateUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framCreateUser 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   100
      TabIndex        =   0
      Top             =   50
      Width           =   5685
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4050
         Left            =   120
         ScaleHeight     =   4050
         ScaleWidth      =   5535
         TabIndex        =   9
         Top             =   150
         Width           =   5535
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            MouseIcon       =   "frmCreateUser.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3600
            Width           =   1455
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   3855
            TabIndex        =   10
            Top             =   2880
            Width           =   3855
            Begin VB.OptionButton optAdmin 
               BackColor       =   &H00FFFFFF&
               Caption         =   "User Administrator"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   2055
            End
            Begin VB.OptionButton optLimited 
               BackColor       =   &H00FFFFFF&
               Caption         =   "User Limited"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2280
               TabIndex        =   6
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "&Create"
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
            Height          =   375
            Left            =   2280
            MouseIcon       =   "frmCreateUser.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txtConfirmPassword 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtNewPassword 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   2
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtNewUserName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   1
            Top             =   480
            Width           =   2775
         End
         Begin VB.CheckBox chkShowPasswordCreateuser 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show password"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   4
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Account Type"
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
            Left            =   120
            TabIndex        =   15
            Top             =   2280
            Width           =   1590
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
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
            Left            =   1080
            TabIndex        =   14
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   1080
            TabIndex        =   13
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter User Name:"
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
            Left            =   1080
            TabIndex        =   12
            Top             =   480
            Width           =   1275
         End
         Begin VB.Image Image1 
            Height          =   555
            Left            =   120
            Picture         =   "frmCreateUser.frx":02B0
            Top             =   2760
            Width           =   720
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5280
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Image Image2 
            Height          =   645
            Left            =   120
            Picture         =   "frmCreateUser.frx":17C2
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Information"
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
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   1800
            X2              =   5280
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line3 
            X1              =   1920
            X2              =   5280
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   2160
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   2160
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   2160
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   2160
            Y1              =   3960
            Y2              =   3960
         End
      End
   End
End
Attribute VB_Name = "frmCreateUser"
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
Dim blnuserexist As Boolean
Dim blntextvalidation As Boolean
Dim blnmaxuser As Boolean

Private Sub chkShowPasswordCreateuser_Click()
If chkShowPasswordCreateuser.Value = 1 Then
    txtNewPassword.PasswordChar = ""
    txtConfirmPassword.PasswordChar = ""
ElseIf chkShowPasswordCreateuser = 0 Then
    txtNewPassword.PasswordChar = "•"
    txtConfirmPassword.PasswordChar = "•"
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
Check_validation
If blntextvalidation = False Then
    Exit Sub
ElseIf blntextvalidation = True Then
    Check_Username
End If

'-----------------------------------------------------
On Err GoTo Err
If blnmaxuser = True Then
    blnmaxuser = False
    MsgBox "You have reached the maximum number of users." & vbCrLf & "No more User Accounts can be created...!", vbExclamation
    txtConfirmPassword = ""
    txtNewPassword = ""
    txtNewUserName = ""
    chkShowPassword.Value = 0
    txtConfirmPassword.Enabled = False
    txtNewPassword.Enabled = False
    txtNewUserName.Enabled = False
    chkShowPassword.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdCreate.Enabled = False
    Exit Sub
End If
If blnuserexist = True Then
    MsgBox "User name is already exist!", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
  Exit Sub
ElseIf blnuserexist = False Then
rstusernames.AddNew
    rstusernames("USER_NAME") = txtNewUserName.Text
    rstusernames("PASSWORD") = Encrypt(txtConfirmPassword)
    If optAdmin.Value = True Then
        rstusernames("TYPE") = "1"
        rstusernames("READ_P") = "1"
        rstusernames("WRITE_P") = "1"
    ElseIf optLimited.Value = True Then
        rstusernames("TYPE") = "0"
        rstusernames("READ_P") = "1"
        rstusernames("WRITE_P") = "0"
    End If
rstusernames.Update
MsgBox "User Account successfully created.", vbInformation
End If
'Me.cmbUsernames.Clear
'frmPassword.Form_Load
txtNewUserName.Text = ""
txtNewPassword.Text = ""
txtConfirmPassword.Text = ""
txtNewUserName.SetFocus
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtNewUserName.SetFocus
End Sub

Public Sub Check_Username()
 Set rstcheckusername = New ADODB.Recordset
     rstcheckusername.CursorLocation = adUseClient
     rstcheckusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & txtNewUserName & "'", dbcon, adOpenStatic, adLockReadOnly
        If rstusernames.RecordCount = 8 Then: blnmaxuser = True: rstcheckusername.Close: Set rstcheckusername = Nothing: Exit Sub
            If rstcheckusername.RecordCount > 0 Then
                blnuserexist = True
            Else
                blnuserexist = False
            End If
rstcheckusername.Close
Set rstcheckusername = Nothing
End Sub

Public Sub Check_validation()
   blntextvalidation = True
If UCase(txtNewUserName) = "ADMINISTRATOR" Or UCase(txtNewUserName) = "USER" Then
    MsgBox "The name you typed is a system built-in name." & vbCrLf & "Please type a different name.", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
    blntextvalidation = False
   Exit Sub
End If
If txtNewUserName = "" Then
    MsgBox "Please type a User name!", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
    blntextvalidation = False
   Exit Sub
End If
If txtNewPassword = "" Then
    MsgBox "Password is required!", vbExclamation
    txtNewPassword.SetFocus
    blntextvalidation = False
Exit Sub
End If
If txtConfirmPassword.Text = "" Then
    MsgBox "Confirm password is required!", vbExclamation
    txtConfirmPassword.SetFocus
    blntextvalidation = False
Exit Sub
End If
If txtNewPassword <> txtConfirmPassword Then
   MsgBox "Password confirmation failed." & vbCrLf & " Please enter passwords again.", vbCritical
   txtNewPassword.Text = ""
   txtConfirmPassword.Text = ""
   txtNewPassword.SetFocus
   blntextvalidation = False
   Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

On Error Resume Next
Set rstusernames = New ADODB.Recordset
    rstusernames.CursorLocation = adUseClient
    rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    optAdmin.Value = True
    'MsgBox rstusernames.RecordCount
If rstusernames.RecordCount = 8 Or rstusernames.RecordCount > 8 Then
    MsgBox "You have reached the maximum number of users." & vbCrLf & "No more User Accounts cannot be created...!", vbExclamation
    txtConfirmPassword.Enabled = False
    txtNewPassword.Enabled = False
    txtNewUserName.Enabled = False
    chkShowPasswordCreateuser.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdCreate.Enabled = False
End If

Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2

If intaccount_type = 0 Then
    txtConfirmPassword.Enabled = False
    txtNewPassword.Enabled = False
    txtNewUserName.Enabled = False
    chkShowPasswordCreateuser.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdCreate.Enabled = False
Else
    cmdCreate.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstusernames.Close
Set rstusernames = Nothing
Set Me.Picture = Nothing
End Sub

