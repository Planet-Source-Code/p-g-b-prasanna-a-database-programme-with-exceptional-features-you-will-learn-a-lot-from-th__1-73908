VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2925
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   855
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loding, please wait..."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   472
      TabIndex        =   0
      Top             =   367
      Width           =   1980
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Private Sub Form_Load()
Set_Round Me
intPaintFlag = 0
End Sub
Private Sub Form_Paint()
'This is how a Splash window really works without using any timer control.
'It displays on the screen until the frmMain is fully loaded.
'Once it is loaded, the Splash window will hide.
If intPaintFlag = 0 Then
    intPaintFlag = 1
    frmSplash.Refresh
    frmMain.Show
    Me.Hide
End If
End Sub

