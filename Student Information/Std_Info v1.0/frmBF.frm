VERSION 5.00
Begin VB.Form frmBF 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for Email Client Application..."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "frmBF.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   4480
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   4320
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "C&ancel"
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
            Left            =   1920
            MouseIcon       =   "frmBF.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   4560
            Width           =   1215
         End
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   0
            MouseIcon       =   "frmBF.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   2280
            Width           =   4215
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
            Left            =   3240
            MouseIcon       =   "frmBF.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   4560
            Width           =   975
         End
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
            Height          =   1665
            Left            =   0
            MouseIcon       =   "frmBF.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   480
            Width           =   4215
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
            Left            =   0
            MouseIcon       =   "frmBF.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   0
            Width           =   4215
         End
         Begin VB.Image Image1 
            Height          =   510
            Left            =   0
            Picture         =   "frmBF.frx":06A6
            Top             =   4440
            Width           =   435
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   4200
            Y1              =   4320
            Y2              =   4320
         End
      End
   End
End
Attribute VB_Name = "frmBF"
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

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
If intbrowseoption = 1 Then
    If Right(File1.Path, 1) = "\" Then
        frmMain.txtImportDblocation = File1.Path & File1.FileName
    Else
        frmMain.txtImportDblocation = File1.Path & "\" & File1.FileName
    End If
    Unload Me
    Exit Sub
ElseIf intbrowseoption = 2 Then
    If Right(File1.Path, 1) = "\" Then
        frmDatabaseSelectionMsg.txtLocal = File1.Path & File1.FileName
    Else
        frmDatabaseSelectionMsg.txtLocal = File1.Path & "\" & File1.FileName
    End If
    
    Unload Me
    Exit Sub
End If

Dim fpath, Link_Path As String
fpath = File1.Path
    If File1.FileName <> "" Then
        If Right(fpath, 1) = "\" Then
            Link_Path = File1.Path & File1.FileName
        Else
            Link_Path = File1.Path & "\" & File1.FileName
        End If
    Else
        
        Exit Sub
    End If
Write_Registry Email_Client_Loc, Link_Path
frmMain.txtEmailClientLocation = Link_Path
strMail_Client = Link_Path
Unload Me
End Sub

Private Sub cmdOk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Dir1_Change()
If intbrowseoption = 1 Or intbrowseoption = 2 Then
    File1.Pattern = "*.mdb"
    File1.Path = Dir1.Path
    If File1.FileName <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    Exit Sub
End If

File1.Pattern = "*.exe"
File1.Path = Dir1.Path
If File1.FileName <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
End Sub

Private Sub Dir1_Click()
If intbrowseoption = 1 Or intbrowseoption = 2 Then
    Dir1.Path = Dir1.Path
    File1.Pattern = "*.mdb"
    File1.Path = Dir1.Path
    If File1.FileName <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    Exit Sub
End If


Dir1.Path = Dir1.Path
File1.Pattern = "*.exe"
File1.Path = Dir1.Path
If File1.FileName <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
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


Private Sub File1_Click()
File1.Path = Dir1.Path
If File1.FileName <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
End Sub

Private Sub File1_DblClick()
cmdOk_Click
End Sub

Private Sub Form_Activate()
On Error Resume Next
cmdOK.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub
Public Sub Err_Path()
Drive1.Drive = UCase(Left(App.Path, 1))
Dir1.Path = App.Path
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

On Error GoTo Err
Dim emclientpath, emclientexe, emclientlocation As String
    If intbrowseoption = 2 Then
        Me.Left = frmDatabaseSelectionMsg.Left + (frmDatabaseSelectionMsg.Width - Me.Width) / 2
        Me.Top = frmDatabaseSelectionMsg.Top + (frmDatabaseSelectionMsg.Height - Me.Height) / 2

    Else
        Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
        Me.Top = frmMain.Top + (frmMain.Height - Me.Height) / 2
    End If
    
    If intbrowseoption = 1 Or intbrowseoption = 2 Then
        Me.Caption = "Select the database..."
        File1.Pattern = "*.mdb"
        File1.Path = Dir1.Path
            If File1.FileName <> "" Then
                cmdOK.Enabled = True
            Else
                cmdOK.Enabled = False
            End If
        Exit Sub
    End If

File1.Pattern = "*.exe"
File1.Path = Dir1.Path
emclientpath = Read_Registry(Email_Client_Loc)
emclientexe = Dir(emclientpath)
emclientlocation = Mid(emclientpath, 1, (Len(emclientpath) - Len(emclientexe)) - 1)
Drive1.Drive = Left(emclientpath, 1)
Dir1.Path = emclientlocation
If File1.FileName <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If

Exit Sub
Err:
Err_Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Me.Picture = Nothing
End Sub
