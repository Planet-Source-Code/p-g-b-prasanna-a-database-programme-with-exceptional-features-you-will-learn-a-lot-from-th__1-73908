VERSION 5.00
Begin VB.Form frmFN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define Folder Name"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   20
      Width           =   3375
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   50
         ScaleHeight     =   945
         ScaleWidth      =   3255
         TabIndex        =   4
         Top             =   120
         Width           =   3255
         Begin VB.TextBox txtFName 
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
            Left            =   80
            MouseIcon       =   "frmFN.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   0
            Top             =   120
            Width           =   3135
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   600
            MouseIcon       =   "frmFN.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   465
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmFN.frx":02A4
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   465
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmFN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Private Sub cmdOk_Click()
On Error GoTo Err_Check
MkDir frmBrowseFolder.Dir1.Path & "\" & txtFName.Text
    If Right(frmBrowseFolder.Dir1.Path, 1) = "\" Then
        frmBrowseFolder.Dir1.Path = frmBrowseFolder.Drive1.Drive & "\" & txtFName.Text
        Unload Me
        frmBrowseFolder.Dir1.Refresh
        Exit Sub
    Else
        frmBrowseFolder.Dir1.Path = frmBrowseFolder.Dir1.Path & "\" & txtFName.Text
        Unload Me
        frmBrowseFolder.Dir1.Refresh
    End If
Exit Sub

Err_Check:
If Err.Number = 76 Then
    MsgBox "Invalid folder name..", vbCritical
    txtFName.Text = ""
    txtFName.SetFocus
ElseIf Err.Number = 75 Then
    MsgBox "Folder already exist..", vbExclamation
    txtFName.SetFocus
    SendKeys "{Home}+{End}"
Else
    MsgBox Err.Number, vbOKOnly
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

Me.Left = frmBrowseFolder.Left + (frmBrowseFolder.Width - Me.Width) / 2
Me.Top = frmBrowseFolder.Top + (frmBrowseFolder.Height - Me.Height) / 2
cmdOK.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Me.Picture = Nothing
End Sub

Private Sub txtFName_Change()
If txtFName.Text <> "" Then
cmdOK.Enabled = True
Else: cmdOK.Enabled = False
End If
End Sub
