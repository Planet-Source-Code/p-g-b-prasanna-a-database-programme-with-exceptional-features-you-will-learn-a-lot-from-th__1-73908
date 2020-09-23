VERSION 5.00
Begin VB.Form frmBrowseFolder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for folders..."
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   Icon            =   "frmBrowseFolder.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Height          =   3855
      Left            =   80
      TabIndex        =   0
      Top             =   50
      Width           =   4720
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   4485
         TabIndex        =   1
         Top             =   120
         Width           =   4480
         Begin VB.DirListBox Dir1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   20
            MouseIcon       =   "frmBrowseFolder.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   840
            Width           =   4455
         End
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   20
            MouseIcon       =   "frmBrowseFolder.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   360
            Width           =   4455
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "&New Folder"
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
            Left            =   2055
            MouseIcon       =   "frmBrowseFolder.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
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
            Left            =   3495
            MouseIcon       =   "frmBrowseFolder.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the folder"
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
            Left            =   20
            TabIndex        =   6
            Top             =   80
            Width           =   2295
         End
         Begin VB.Line Line1 
            X1              =   15
            X2              =   4455
            Y1              =   3100
            Y2              =   3100
         End
         Begin VB.Image Image1 
            Height          =   495
            Left            =   0
            Picture         =   "frmBrowseFolder.frx":0554
            Top             =   3120
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Public ret As String
Public int_click As Boolean
Private Sub cmdNewFolder_Click()
frmFN.Show 1
End Sub

Private Sub cmdNewFolder_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub
Private Sub cmdOk_Click()
Dim retfs As String
'On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
    ret = Dir1.Path
    retfs = UCase(Left(ret, 1)) + Mid(ret, 2, Len(ret) - 1)
    Write_Registry (Default_File_Loc), retfs
    frmMain.txtDefaultFileSavingLocation = retfs
    strFile_Save = retfs
    Unload Me
            
Else
    ret = Dir1.Path & "\"
    retfs = UCase(Left(ret, 1)) + Mid(ret, 2, Len(ret) - 1)
    Write_Registry (Default_File_Loc), retfs
    frmMain.txtDefaultFileSavingLocation = retfs
    strFile_Save = retfs
    Unload Me
End If
End Sub

Private Sub cmdOk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Dir1_Click()
Dir1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo Err_Check
Dir1.Path = Drive1.Drive
Exit Sub
Err_Check:
MsgBox "Device is not ready..!", vbCritical
Drive1.Refresh
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Activate()
cmdOK.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

On Error GoTo Err
Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2
Drive1.Drive = Left(Read_Registry(Default_File_Loc), 1)
Dir1.Path = Read_Registry(Default_File_Loc)
'Check_for_Default_Path
'If Folder_Exist = False Then
'Err_Path
'End If
Exit Sub
Err:
Err_Path
End Sub

Public Sub Err_Path()
Drive1.Drive = UCase(Left(App.Path, 1))
Dir1.Path = App.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Me.Picture = Nothing
End Sub
