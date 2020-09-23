VERSION 5.00
Begin VB.Form frmManageUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage User..."
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmManageUser.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framManageUser 
      BackColor       =   &H00FFFFFF&
      Height          =   8700
      Left            =   100
      TabIndex        =   0
      Top             =   80
      Width           =   6030
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8445
         Left            =   120
         ScaleHeight     =   8445
         ScaleWidth      =   5775
         TabIndex        =   15
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "Close"
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
            Left            =   3720
            MouseIcon       =   "frmManageUser.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   8040
            Width           =   1815
         End
         Begin VB.ComboBox cmbUsernames 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   3495
         End
         Begin VB.CommandButton cmdDeleteAccount 
            Caption         =   "Delete"
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
            Left            =   2955
            MouseIcon       =   "frmManageUser.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox chkShowPassword 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   3645
            Width           =   1695
         End
         Begin VB.CommandButton cmdChangePrivileges 
            Caption         =   "Change Privileges"
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
            Left            =   2880
            MouseIcon       =   "frmManageUser.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   7200
            Width           =   2655
         End
         Begin VB.CheckBox chkWritePermission 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Write Permission"
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
            Left            =   1440
            TabIndex        =   13
            Top             =   6840
            Width           =   4095
         End
         Begin VB.CheckBox chkReadPermission 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Read Permission"
            Enabled         =   0   'False
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
            Left            =   1440
            TabIndex        =   12
            Top             =   6480
            Width           =   4095
         End
         Begin VB.CommandButton cmdChangeUserType 
            Caption         =   "Change User Type"
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
            Left            =   2955
            MouseIcon       =   "frmManageUser.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   5520
            Width           =   2535
         End
         Begin VB.OptionButton optLimited 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Limited"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   10
            Top             =   5040
            Width           =   1335
         End
         Begin VB.OptionButton optAdmin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Administrator"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   9
            Top             =   5040
            Width           =   1935
         End
         Begin VB.TextBox txtOldPassword 
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
            Left            =   2835
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   2280
            Width           =   2655
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
            Left            =   2835
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   4
            Top             =   2640
            Width           =   2655
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
            Left            =   2835
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   5
            Top             =   3000
            Width           =   2655
         End
         Begin VB.CommandButton cmdChangePassword 
            Caption         =   "Change Password"
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
            Left            =   2955
            MouseIcon       =   "frmManageUser.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3600
            Width           =   2535
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   5520
            Y1              =   7920
            Y2              =   7920
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   3600
            Y1              =   8160
            Y2              =   8160
         End
         Begin VB.Line Line6 
            X1              =   240
            X2              =   3600
            Y1              =   8280
            Y2              =   8280
         End
         Begin VB.Line Line8 
            X1              =   240
            X2              =   3600
            Y1              =   8400
            Y2              =   8400
         End
         Begin VB.Line Line9 
            X1              =   240
            X2              =   3600
            Y1              =   8040
            Y2              =   8040
         End
         Begin VB.Image Image5 
            Height          =   135
            Left            =   120
            Picture         =   "frmManageUser.frx":06A6
            Top             =   600
            Width           =   5550
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select User Name"
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
            Left            =   360
            TabIndex        =   23
            Top             =   240
            Width           =   1485
         End
         Begin VB.Image Image4 
            Height          =   720
            Left            =   360
            Picture         =   "frmManageUser.frx":2E00
            Top             =   1080
            Width           =   735
         End
         Begin VB.Line Line7 
            X1              =   2160
            X2              =   5520
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete User"
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
            Left            =   360
            TabIndex        =   22
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblResetPassword 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reset User Password..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3600
            MouseIcon       =   "frmManageUser.frx":4A02
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   4200
            Width           =   1920
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "User Action Privileges"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   6120
            Width           =   2055
         End
         Begin VB.Line Line3 
            X1              =   2400
            X2              =   5520
            Y1              =   6240
            Y2              =   6240
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   360
            Picture         =   "frmManageUser.frx":4B54
            Top             =   6480
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   750
            Left            =   360
            Picture         =   "frmManageUser.frx":6816
            Top             =   5040
            Width           =   780
         End
         Begin VB.Image Image2 
            Height          =   990
            Left            =   240
            Picture         =   "frmManageUser.frx":86D0
            Top             =   2520
            Width           =   735
         End
         Begin VB.Line Line2 
            X1              =   2160
            X2              =   5520
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Change User Type"
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
            Left            =   360
            TabIndex        =   20
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Line Line1 
            X1              =   2160
            X2              =   5520
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Change Password"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old Password"
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
            TabIndex        =   18
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
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
            TabIndex        =   17
            Top             =   2640
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm New Password"
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
            TabIndex        =   16
            Top             =   3000
            Width           =   1650
         End
      End
   End
End
Attribute VB_Name = "frmManageUser"
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
Dim oldpass As String
Private Sub chkShowPassword_Click()
If chkShowPassword.Value = 1 Then
    txtOldPassword.PasswordChar = ""
    txtNewPassword.PasswordChar = ""
    txtConfirmPassword.PasswordChar = ""
ElseIf chkShowPassword.Value = 0 Then
    txtOldPassword.PasswordChar = "•"
    txtNewPassword.PasswordChar = "•"
    txtConfirmPassword.PasswordChar = "•"
End If
End Sub
Private Sub cmbUsernames_Click()
Commands_Set
End Sub

Private Sub cmdChangePassword_Click()
Password_Change_Pro
End Sub

Private Sub cmdChangePrivileges_Click()
User_Privileges_Change_Pro
End Sub

Private Sub cmdChangeUserType_Click()
User_Type_Change_Pro
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeleteAccount_Click()
If cmbUsernames = User Then
    MsgBox "You can't delete the current account." & vbCrLf & "Please Login with another Administrator Account.", vbCritical
    Exit Sub
End If
 Set rstusernamefordelete = New ADODB.Recordset
     rstusernamefordelete.CursorLocation = adUseClient
      rstusernamefordelete.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames.Text & "'", dbcon, adOpenStatic, adLockOptimistic
If MsgBox("Are you sure you need to delete the account " & cmbUsernames & " ?", vbYesNo + vbQuestion) = vbYes Then
    If rstusernamefordelete.RecordCount > 0 Then
        If UCase(rstusernamefordelete("USER_NAME")) = UCase("Administrator") Then
            MsgBox "The Account " & cmbUsernames & " is a built-in account." & vbCrLf & "It cannot be deleted.", vbCritical
            Exit Sub
        End If
            On Error GoTo Err
            rstusernamefordelete.Delete
            MsgBox "User Account " & cmbUsernames & " successfully deleted.", vbInformation
            cmbUsernames.Clear
            Form_Load
        End If
End If
rstusernamefordelete.Close
Set rstusernamefordelete = Nothing
Exit Sub

Err:
    MsgBox Err.Description & " _ " & Err.Number & "."
    Unload Me
End Sub

'Private Sub cmdClose_KeyPress(KeyAscii As Integer)
'If KeyAscii = 162 Then
'    PWORD_INFO
'End If
'End Sub

Private Sub Form_Activate()
On Error Resume Next
txtOldPassword.SetFocus
If cmbUsernames <> User Then
    lblResetPassword.Enabled = True
Else
    lblResetPassword.Enabled = False
End If

'MsgBox " left " & frmChangePassword.Left & vbCrLf & "top - " & frmChangePassword.Top
End Sub
Public Sub Add_User_Names_to_Combo()
If rstusernames.RecordCount > 0 Then
    cmbUsernames.Clear
    Do While Not rstusernames.EOF
        cmbUsernames.AddItem rstusernames("USER_NAME")
        rstusernames.MoveNext
    Loop
End If
rstusernames.Close
Set rstusernames = Nothing
End Sub

Public Sub Password_Change_Pro()
 Set rstgetusername = New ADODB.Recordset
     rstgetusername.CursorLocation = adUseClient
     rstgetusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
      oldpass = Decrypt(rstgetusername("PASSWORD"))
      
     If txtOldPassword <> oldpass Then
        MsgBox "Old password is wrong.", vbCritical
            With txtOldPassword
                .SetFocus
                .SelStart = 0
                .SelLength = Len(txtOldPassword)
            End With
        Exit Sub
     End If
     
     If txtNewPassword.Text = "" Then
        MsgBox "Password is required.", vbExclamation
            With txtNewPassword
                .SetFocus
                .SelStart = 0
                .SelLength = Len(txtNewPassword)
            End With
        Exit Sub
     End If
     
     If txtConfirmPassword.Text = "" Then
        MsgBox "Confirm password is required.", vbExclamation
            With txtConfirmPassword
                .SetFocus
                .SelStart = 0
                .SelLength = Len(txtConfirmPassword)
            End With
        Exit Sub
     End If
     
     If txtNewPassword <> txtConfirmPassword Then
        MsgBox "Password confirmation failed." & vbCrLf & "Please enter passwords again.", vbCritical
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
            With txtNewPassword
                .SetFocus
                .SelStart = 0
                .SelLength = Len(txtNewPassword)
            End With
        Exit Sub
     ElseIf txtNewPassword = txtConfirmPassword Then
        rstgetusername("PASSWORD") = Encrypt(txtConfirmPassword)
        rstgetusername.Update
        MsgBox "Password successfully changed." & vbCrLf & "Log in again for the changes.", vbInformation
        txtOldPassword.Text = ""
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
        txtOldPassword.SetFocus
     End If
rstgetusername.Close
Set rstgetusername = Nothing
End Sub

Public Sub User_Type_Change_Pro()
On Error GoTo Err
 Set rstgetusername = New ADODB.Recordset
 rstgetusername.CursorLocation = adUseClient
 rstgetusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
     If optAdmin.Value = True Then
         rstgetusername("TYPE") = "1"
         rstgetusername.Update
     ElseIf optLimited.Value = True Then
         rstgetusername("TYPE") = "0"
         rstgetusername.Update
     End If
MsgBox "User Type successfully changed." & vbCrLf & "Log In again for the changes." & _
vbCrLf & vbCrLf & "Tip" & vbCrLf & "---" & vbCrLf & "You may need to change the Privileges as well.", vbInformation
rstgetusername.Close
Set rstgetusername = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub
Public Sub User_Privileges_Change_Pro()
On Error GoTo Err
Set rstuserforprivileges = New ADODB.Recordset
rstuserforprivileges.CursorLocation = adUseClient
rstuserforprivileges.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
 
rstuserforprivileges("READ_P") = chkReadPermission.Value
rstuserforprivileges("WRITE_P") = chkWritePermission.Value
     
rstuserforprivileges.Update
MsgBox "User Privileges successfully changed.", vbInformation
        
rstuserforprivileges.Close
Set rstuserforprivileges = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub
Public Sub Account_type_initialize()
If intaccount_type = 0 Then
    cmbUsernames.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdChangeUserType.Enabled = False
    cmdDeleteAccount.Enabled = False
End If
cmbUsernames = User
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

Set rstusernames = New ADODB.Recordset
    rstusernames.CursorLocation = adUseClient
    rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
Add_User_Names_to_Combo

On Error Resume Next
cmbUsernames.Text = cmbUsernames.List(0)
Account_type_initialize

Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Me.Picture = Nothing
End Sub

Private Sub lblResetPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.Top = lblResetPassword.Top + 20
lblResetPassword.Left = lblResetPassword.Left + 20

End Sub

Private Sub lblResetPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.ForeColor = &HC0&
End Sub

Private Sub lblResetPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.Top = lblResetPassword.Top - 20
lblResetPassword.Left = lblResetPassword.Left - 20

On Error GoTo Err
Set rstresetpassword = New ADODB.Recordset
rstresetpassword.CursorLocation = adUseClient
rstresetpassword.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
    If MsgBox("Are you sure you want to reset the password for " & cmbUsernames & " ?", vbQuestion + vbYesNo) = vbYes Then
        rstresetpassword("PASSWORD") = Encrypt("password")
        rstresetpassword.Update
        MsgBox "The password of " & cmbUsernames & " has been reset as 'password' ", vbInformation
    End If

rstresetpassword.Close
Set rstresetpassword = Nothing
Exit Sub

Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'unload Me
Form_Load
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.ForeColor = &HC00000
End Sub

Public Sub Commands_Set()
 On Error Resume Next
 Set rstchecktype = New ADODB.Recordset
 rstchecktype.CursorLocation = adUseClient
 rstchecktype.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockReadOnly
     If rstchecktype("TYPE") = "1" Then
         optAdmin.Value = True
     ElseIf rstchecktype("TYPE") = "0" Then
         optLimited.Value = True
     End If
     
     If rstchecktype("READ_P") = "1" Then
         chkReadPermission.Value = 1
     Else
         chkReadPermission.Value = 0
     End If
     
     If rstchecktype("WRITE_P") = "1" Then
         chkWritePermission.Value = 1
     Else
         chkWritePermission.Value = 0
     End If
     
     If intaccount_type = 0 Then
            cmdDeleteAccount.Enabled = False
     Else
            cmdDeleteAccount.Enabled = True
     End If
      
 If cmbUsernames = "Administrator" Then
     'optAdmin.Enabled = False
     optLimited.Enabled = False: cmdChangeUserType.Enabled = False
     lblResetPassword.Enabled = False: chkReadPermission.Enabled = False
     chkWritePermission.Enabled = False
         If User <> "Administrator" Then
              txtOldPassword.Enabled = False: txtNewPassword.Enabled = False
              txtConfirmPassword.Enabled = False: chkShowPassword.Enabled = False
              chkShowPassword.Value = 0: cmdChangePassword.Enabled = False
         ElseIf User = "Administrator" Then
              txtOldPassword.Enabled = True: txtNewPassword.Enabled = True
              txtConfirmPassword.Enabled = True: chkShowPassword.Value = 0
              chkShowPassword.Enabled = True: cmdChangePassword.Enabled = True
              cmdChangePrivileges.Enabled = False
         End If
 Else
     If cmbUsernames = User Then
         txtOldPassword.Enabled = True: txtNewPassword.Enabled = True
         txtConfirmPassword.Enabled = True: chkShowPassword.Value = 0
         chkShowPassword.Enabled = True: cmdChangePassword.Enabled = True
         lblResetPassword.Enabled = False: optAdmin.Enabled = False
         optLimited.Enabled = False: cmdChangeUserType.Enabled = False
         cmdChangePrivileges.Enabled = False: chkReadPermission.Enabled = False
         chkWritePermission.Enabled = False
     Else
         txtOldPassword.Enabled = False: txtNewPassword.Enabled = False
         txtConfirmPassword.Enabled = False: chkShowPassword.Value = 0
         chkShowPassword.Enabled = False: cmdChangePassword.Enabled = False
         lblResetPassword.Enabled = True: optAdmin.Enabled = True
         optLimited.Enabled = True: cmdChangeUserType.Enabled = True
         cmdChangePrivileges.Enabled = True: chkReadPermission.Enabled = True
         chkWritePermission.Enabled = True
     End If
 End If
 chkReadPermission.Enabled = False
 rstchecktype.Close
 Set rstchecktype = Nothing
End Sub
