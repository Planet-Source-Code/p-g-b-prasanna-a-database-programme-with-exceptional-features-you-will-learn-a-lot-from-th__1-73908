VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Information Basis v 1.0 - Singapore Informatics Computer Institute (pvt) Ltd. ©  Bandula"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12750
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   12750
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      TabIndex        =   153
      Top             =   120
      Width           =   9135
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   40
         ScaleHeight     =   450
         ScaleWidth      =   9030
         TabIndex        =   154
         Top             =   120
         Width           =   9030
         Begin VB.Label lblCurrentUser 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   7245
            TabIndex        =   157
            Top             =   120
            Width           =   45
         End
         Begin VB.Label lblUserInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active User:"
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
            Left            =   6120
            TabIndex        =   156
            Top             =   120
            Width           =   1020
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   100
            Picture         =   "frmMain.frx":617A
            Top             =   40
            Width           =   360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Course Information."
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
            Left            =   600
            TabIndex        =   155
            Top             =   105
            Width           =   1890
         End
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   10500
         Picture         =   "frmMain.frx":68E4
         ToolTipText     =   "Application Help"
         Top             =   165
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   9735
         Left            =   40
         ScaleHeight     =   9735
         ScaleWidth      =   3090
         TabIndex        =   66
         Top             =   120
         Width           =   3090
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -1000
            TabIndex        =   91
            Top             =   5880
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -1000
            TabIndex        =   90
            Top             =   5520
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblLogOff 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Log Off"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   675
            MouseIcon       =   "frmMain.frx":704E
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   5760
            Width           =   1305
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   120
            MouseIcon       =   "frmMain.frx":71A0
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":72F2
            Top             =   5640
            Width           =   540
         End
         Begin VB.Image Image9 
            Height          =   585
            Left            =   120
            MouseIcon       =   "frmMain.frx":80B4
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":8206
            Top             =   4440
            Width           =   570
         End
         Begin VB.Image Image10 
            Height          =   495
            Left            =   120
            MouseIcon       =   "frmMain.frx":93F4
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":9546
            Top             =   6360
            Width           =   540
         End
         Begin VB.Image Image18 
            Height          =   660
            Left            =   120
            MouseIcon       =   "frmMain.frx":A374
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":A4C6
            Top             =   3600
            Width           =   690
         End
         Begin VB.Image Image17 
            Height          =   495
            Left            =   120
            MouseIcon       =   "frmMain.frx":BD18
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":BE6A
            Top             =   1080
            Width           =   630
         End
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "frmMain.frx":CF2C
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":D07E
            Top             =   2760
            Width           =   675
         End
         Begin VB.Image Image13 
            Height          =   615
            Left            =   120
            MouseIcon       =   "frmMain.frx":E8A8
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":E9FA
            Top             =   1920
            Width           =   675
         End
         Begin VB.Image Image12 
            Height          =   555
            Left            =   120
            MouseIcon       =   "frmMain.frx":10004
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":10156
            Top             =   240
            Width           =   675
         End
         Begin VB.Image Image11 
            Height          =   570
            Left            =   225
            Picture         =   "frmMain.frx":11540
            Top             =   8040
            Width           =   2685
         End
         Begin VB.Label lblViewExport 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Data View/Export"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   660
            LinkTimeout     =   100
            MouseIcon       =   "frmMain.frx":165AA
            MousePointer    =   99  'Custom
            TabIndex        =   97
            Top             =   360
            Width           =   2370
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000000&
            X1              =   120
            X2              =   3000
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pgbsoft@gmail.com"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   720
            MouseIcon       =   "frmMain.frx":166FC
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   9240
            Width           =   1665
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programmed by Bandula on a request."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   165
            TabIndex        =   74
            Top             =   8880
            Width           =   2835
         End
         Begin VB.Label lblCourseInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   840
            MouseIcon       =   "frmMain.frx":1684E
            MousePointer    =   99  'Custom
            TabIndex        =   72
            Top             =   3720
            Width           =   1400
         End
         Begin VB.Label lblGetEmails 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Get Emails"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   780
            LinkTimeout     =   100
            MouseIcon       =   "frmMain.frx":169A0
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   1200
            Width           =   1380
         End
         Begin VB.Label lblGetContacts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Get Contacts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   825
            MouseIcon       =   "frmMain.frx":16AF2
            MousePointer    =   99  'Custom
            TabIndex        =   70
            Top             =   2040
            Width           =   1580
         End
         Begin VB.Label lblExit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   700
            MouseIcon       =   "frmMain.frx":16C44
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   6480
            Width           =   880
         End
         Begin VB.Label lblStudentInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Student Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   720
            MouseIcon       =   "frmMain.frx":16D96
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   2880
            Width           =   1750
         End
         Begin VB.Label lblOptions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   720
            MouseIcon       =   "frmMain.frx":16EE8
            MousePointer    =   99  'Custom
            TabIndex        =   67
            Top             =   4560
            Width           =   1250
         End
      End
   End
   Begin VB.Frame framOptions 
      BackColor       =   &H00FFFFFF&
      Height          =   9360
      Left            =   3480
      TabIndex        =   73
      Top             =   740
      Width           =   9135
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   650
         Left            =   120
         ScaleHeight     =   645
         ScaleWidth      =   8775
         TabIndex        =   168
         Top             =   120
         Width           =   8775
         Begin VB.Image Image14 
            Height          =   540
            Left            =   0
            MouseIcon       =   "frmMain.frx":1703A
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":1718C
            Top             =   45
            Width           =   525
         End
         Begin VB.Label lblCreateUser 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Create User"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   420
            MouseIcon       =   "frmMain.frx":180FE
            MousePointer    =   99  'Custom
            TabIndex        =   171
            Top             =   160
            Width           =   1515
         End
         Begin VB.Image Image15 
            Height          =   600
            Left            =   2160
            MouseIcon       =   "frmMain.frx":18250
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":183A2
            Top             =   40
            Width           =   525
         End
         Begin VB.Label lblManageUser 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Manage User"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            MouseIcon       =   "frmMain.frx":194C4
            MousePointer    =   99  'Custom
            TabIndex        =   170
            Top             =   160
            Width           =   1620
         End
         Begin VB.Image Image19 
            Height          =   525
            Left            =   4560
            MouseIcon       =   "frmMain.frx":19616
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":19768
            Top             =   40
            Width           =   495
         End
         Begin VB.Label lblBackupImport 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Backup/Import"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   5040
            MouseIcon       =   "frmMain.frx":1A556
            MousePointer    =   99  'Custom
            TabIndex        =   169
            Top             =   160
            Width           =   1815
         End
      End
      Begin VB.Frame framGeneral 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8480
         Left            =   50
         TabIndex        =   123
         Top             =   840
         Width           =   9050
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   8300
            Left            =   40
            ScaleHeight     =   8295
            ScaleWidth      =   8940
            TabIndex        =   124
            Top             =   120
            Width           =   8940
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   140
               Left            =   4770
               ScaleHeight     =   135
               ScaleWidth      =   4005
               TabIndex        =   163
               Top             =   7980
               Width           =   4000
               Begin VB.Label lblProgressImport1 
                  BackColor       =   &H001BD805&
                  Height          =   80
                  Left            =   0
                  TabIndex        =   165
                  Top             =   45
                  Width           =   15
               End
               Begin VB.Label lblProgressImport2 
                  BackColor       =   &H00B3F9B8&
                  Height          =   45
                  Left            =   0
                  TabIndex        =   164
                  Top             =   0
                  Width           =   15
               End
            End
            Begin VB.CommandButton cmdImportData 
               Caption         =   "&Import Data"
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
               Left            =   7440
               MouseIcon       =   "frmMain.frx":1A6A8
               MousePointer    =   99  'Custom
               TabIndex        =   139
               Top             =   7400
               Width           =   1335
            End
            Begin VB.CommandButton cmdSelectDatabase 
               Caption         =   "Select..."
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
               Left            =   6000
               MouseIcon       =   "frmMain.frx":1A7FA
               MousePointer    =   99  'Custom
               TabIndex        =   138
               Top             =   7400
               Width           =   1335
            End
            Begin VB.TextBox txtImportDblocation 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   137
               Top             =   7440
               Width           =   4095
            End
            Begin VB.CommandButton cmdFormatDatabase 
               Caption         =   "Format Database"
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
               Left            =   6120
               MouseIcon       =   "frmMain.frx":1A94C
               MousePointer    =   99  'Custom
               TabIndex        =   136
               Top             =   6300
               Width           =   1695
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear..."
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
               Left            =   6120
               MouseIcon       =   "frmMain.frx":1AA9E
               MousePointer    =   99  'Custom
               TabIndex        =   135
               Top             =   5720
               Width           =   1695
            End
            Begin VB.ComboBox cmbClearDatabase 
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
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   5760
               Width           =   4095
            End
            Begin VB.ComboBox cmbEmailSeperator 
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
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   4320
               Width           =   4095
            End
            Begin VB.CommandButton cmdSetasDefaultEmSeperator 
               Caption         =   "Set as default..."
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
               Left            =   6120
               MouseIcon       =   "frmMain.frx":1ABF0
               MousePointer    =   99  'Custom
               TabIndex        =   132
               Top             =   4300
               Width           =   1695
            End
            Begin VB.CommandButton cmdSetasdefault 
               Caption         =   "Set as default..."
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
               Left            =   6120
               MouseIcon       =   "frmMain.frx":1AD42
               MousePointer    =   99  'Custom
               TabIndex        =   131
               Top             =   3460
               Width           =   1695
            End
            Begin VB.ComboBox cmbLoadPrompt 
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
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   130
               Top             =   3480
               Width           =   4095
            End
            Begin VB.CheckBox chkOpenFile 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Automatically open files after file saving is done (Recommended)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1800
               TabIndex        =   129
               Top             =   2200
               Width           =   6975
            End
            Begin VB.TextBox txtDefaultFileSavingLocation 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   128
               Top             =   1560
               Width           =   5415
            End
            Begin VB.CommandButton cmdDefaultFLocation 
               Caption         =   "Browse..."
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
               Left            =   7440
               MouseIcon       =   "frmMain.frx":1AE94
               MousePointer    =   99  'Custom
               TabIndex        =   127
               Top             =   1530
               Width           =   1335
            End
            Begin VB.TextBox txtEmailClientLocation 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   126
               Top             =   720
               Width           =   5415
            End
            Begin VB.CommandButton SelectEMclient 
               Caption         =   "Browse..."
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
               Left            =   7440
               MouseIcon       =   "frmMain.frx":1AFE6
               MousePointer    =   99  'Custom
               TabIndex        =   125
               Top             =   690
               Width           =   1335
            End
            Begin VB.Label lblGoToLocation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Go to file saving location..."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   165
               Left            =   5560
               MouseIcon       =   "frmMain.frx":1B138
               MousePointer    =   99  'Custom
               TabIndex        =   183
               Top             =   1920
               Width           =   1635
            End
            Begin VB.Label lblImportDescription 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importing data, please wait..."
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   2640
               TabIndex        =   167
               Top             =   7965
               Visible         =   0   'False
               Width           =   1710
            End
            Begin VB.Label lblImportIndicator 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "100%"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   165
               Left            =   4410
               TabIndex        =   166
               Top             =   7970
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Line Line10 
               X1              =   2880
               X2              =   8760
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Programme Path Configuration"
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
               TabIndex        =   151
               Top             =   0
               Width           =   2625
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clear Database Location..."
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   165
               Left            =   1800
               MouseIcon       =   "frmMain.frx":1B28A
               MousePointer    =   99  'Custom
               TabIndex        =   150
               Top             =   6720
               Width           =   1635
            End
            Begin VB.Image Image20 
               Height          =   1320
               Left            =   120
               Picture         =   "frmMain.frx":1B3DC
               Top             =   3240
               Width           =   1395
            End
            Begin VB.Image Image21 
               Height          =   1230
               Left            =   120
               Picture         =   "frmMain.frx":2145E
               Top             =   840
               Width           =   1380
            End
            Begin VB.Image Image22 
               Height          =   1335
               Left            =   120
               Picture         =   "frmMain.frx":26D08
               Top             =   6000
               Width           =   1335
            End
            Begin VB.Line Line9 
               X1              =   2280
               X2              =   8760
               Y1              =   5040
               Y2              =   5040
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Database Configuration"
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
               TabIndex        =   149
               Top             =   4920
               Width           =   1995
            End
            Begin VB.Line Line8 
               X1              =   2680
               X2              =   8760
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Programme Default Settings"
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
               TabIndex        =   148
               Top             =   2640
               Width           =   2415
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Import data from a same structured database:"
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
               Left            =   1800
               TabIndex        =   147
               Top             =   7080
               Width           =   3360
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Format database for data export (Recommended):"
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
               Left            =   1800
               TabIndex        =   146
               Top             =   6360
               Width           =   3660
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clear Database:"
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
               Left            =   1800
               TabIndex        =   145
               Top             =   5400
               Width           =   1170
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Default e-mail address seperator:"
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
               Left            =   1800
               TabIndex        =   144
               Top             =   3960
               Width           =   2415
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Programe default focus at start:"
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
               Left            =   1800
               TabIndex        =   143
               Top             =   3120
               Width           =   2325
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Default File Saving Location:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1800
               TabIndex        =   142
               Top             =   1200
               Width           =   2445
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Email Client Location (Eg: Location of Outlook Express):"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1800
               TabIndex        =   141
               Top             =   360
               Width           =   5385
            End
            Begin VB.Label Label51 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00472BEA&
               Height          =   195
               Left            =   3720
               TabIndex        =   140
               Top             =   2925
               Width           =   945
            End
         End
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   8880
         Y1              =   800
         Y2              =   800
      End
   End
   Begin VB.Frame framStudentInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   9360
      Left            =   3480
      TabIndex        =   48
      Top             =   740
      Width           =   9135
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8235
         Left            =   120
         ScaleHeight     =   8235
         ScaleWidth      =   8940
         TabIndex        =   49
         Top             =   240
         Width           =   8940
         Begin VB.ComboBox cmbSearchOption 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtDateofBirth 
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
            Left            =   6600
            TabIndex        =   23
            Top             =   4200
            Width           =   2175
         End
         Begin VB.ComboBox cmbSex 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   4200
            Width           =   2895
         End
         Begin VB.TextBox txtNICPassport 
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
            Left            =   6600
            TabIndex        =   20
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox txtExamIndex 
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
            Left            =   6600
            TabIndex        =   18
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CommandButton cmdCancelSearch 
            Caption         =   "Cancel"
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
            Height          =   375
            Left            =   7680
            MouseIcon       =   "frmMain.frx":2CA76
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox txtSearch 
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
            Left            =   3240
            TabIndex        =   14
            Top             =   720
            Width           =   3135
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
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
            Left            =   6480
            MouseIcon       =   "frmMain.frx":2CBC8
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox txtRemarks 
            BackColor       =   &H00FFFFFF&
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
            Height          =   525
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   6345
            Width           =   5055
         End
         Begin VB.TextBox txtStdTel 
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
            Left            =   1920
            TabIndex        =   25
            Top             =   5280
            Width           =   4455
         End
         Begin VB.TextBox txtStdEmail 
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
            Left            =   1920
            TabIndex        =   26
            Top             =   5805
            Width           =   4455
         End
         Begin VB.ComboBox cmbCourse 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox txtStdAddress 
            BackColor       =   &H00FFFFFF&
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
            Height          =   525
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   4635
            Width           =   3975
         End
         Begin VB.CommandButton cmdStdAdd 
            Caption         =   "&Add"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":2CD1A
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   7320
            Width           =   2055
         End
         Begin VB.CommandButton cmdStdEdit 
            Caption         =   "&Edit"
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
            Left            =   2280
            MouseIcon       =   "frmMain.frx":2CE6C
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   7320
            Width           =   1575
         End
         Begin VB.CommandButton cmdStdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7680
            MouseIcon       =   "frmMain.frx":2CFBE
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   7320
            Width           =   1095
         End
         Begin VB.CommandButton cmdStdDelete 
            Caption         =   "&Delete"
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
            Left            =   3960
            MouseIcon       =   "frmMain.frx":2D110
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   7320
            Width           =   1935
         End
         Begin VB.CommandButton cmdStdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
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
            Left            =   6000
            MouseIcon       =   "frmMain.frx":2D262
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   7320
            Width           =   1575
         End
         Begin VB.CommandButton cmdStdNext 
            Caption         =   "&Next >"
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
            Left            =   3960
            MouseIcon       =   "frmMain.frx":2D3B4
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   7800
            Width           =   1935
         End
         Begin VB.CommandButton cmdStdPrevious 
            Caption         =   "< &Previous"
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
            Left            =   2280
            MouseIcon       =   "frmMain.frx":2D506
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   7800
            Width           =   1575
         End
         Begin VB.CommandButton cmdStdFirst 
            Caption         =   "|< &First"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":2D658
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   7800
            Width           =   2055
         End
         Begin VB.CommandButton cmdStdLast 
            Caption         =   "&Last >|"
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
            Left            =   6000
            MouseIcon       =   "frmMain.frx":2D7AA
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   7800
            Width           =   1575
         End
         Begin VB.TextBox txtStdName 
            BackColor       =   &H00FFFFFF&
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
            Height          =   525
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   3480
            Width           =   3975
         End
         Begin VB.TextBox txtStdID 
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
            Left            =   1920
            TabIndex        =   17
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Parameter:"
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
            Left            =   3240
            TabIndex        =   89
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(dd/mm/yyyy)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   6600
            TabIndex        =   88
            Top             =   4680
            Width           =   1020
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Option:"
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
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   1080
         End
         Begin VB.Image Image8 
            Height          =   720
            Left            =   7560
            Picture         =   "frmMain.frx":2D8FC
            Top             =   5400
            Width           =   645
         End
         Begin VB.Line Line6 
            X1              =   1560
            X2              =   8760
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Search Student"
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
            Left            =   120
            TabIndex        =   86
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth:"
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
            Left            =   5280
            TabIndex        =   85
            Top             =   4260
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
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
            Left            =   120
            TabIndex        =   84
            Top             =   4200
            Width           =   330
         End
         Begin VB.Line Line5 
            X1              =   2760
            X2              =   8760
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Student Personal Information"
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
            Left            =   120
            TabIndex        =   83
            Top             =   3000
            Width           =   3015
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Official Information"
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
            Left            =   120
            TabIndex        =   82
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIC/Passport No:"
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
            Left            =   5040
            TabIndex        =   81
            Top             =   2400
            Width           =   1260
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Examination Index:"
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
            Left            =   5040
            TabIndex        =   80
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   8760
            Y1              =   7080
            Y2              =   7080
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   2880
            TabIndex        =   57
            Top             =   120
            Width           =   1425
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            X1              =   1920
            X2              =   8760
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Left            =   120
            TabIndex        =   56
            Top             =   6480
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
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
            Left            =   120
            TabIndex        =   55
            Top             =   5880
            Width           =   420
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number:"
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
            Left            =   120
            TabIndex        =   54
            Top             =   5355
            Width           =   1230
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Course:"
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
            Left            =   120
            TabIndex        =   53
            Top             =   2400
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Address:"
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
            Left            =   120
            TabIndex        =   52
            Top             =   4800
            Width           =   1260
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Name:"
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
            Left            =   120
            TabIndex        =   51
            Top             =   3600
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student ID:"
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
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   840
         End
      End
   End
   Begin VB.Frame framCourseInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   9360
      Left            =   3480
      TabIndex        =   44
      Top             =   740
      Width           =   9135
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8355
         Left            =   120
         ScaleHeight     =   8355
         ScaleWidth      =   8940
         TabIndex        =   10
         Top             =   240
         Width           =   8940
         Begin VB.TextBox txtCourse 
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
            Left            =   1920
            TabIndex        =   1
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtCourseDes 
            BackColor       =   &H00FFFFFF&
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
            Height          =   525
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   840
            Width           =   6855
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "&Last >|"
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
            Left            =   6000
            MouseIcon       =   "frmMain.frx":2F1FE
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "|< &First"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":2F350
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "< &Previous"
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
            MouseIcon       =   "frmMain.frx":2F4A2
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next >"
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
            Left            =   4080
            MouseIcon       =   "frmMain.frx":2F5F4
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
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
            Left            =   6000
            MouseIcon       =   "frmMain.frx":2F746
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            Left            =   4080
            MouseIcon       =   "frmMain.frx":2F898
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7680
            MouseIcon       =   "frmMain.frx":2F9EA
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
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
            MouseIcon       =   "frmMain.frx":2FB3C
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":2FC8E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   8760
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Image Image3 
            Height          =   720
            Left            =   7920
            Picture         =   "frmMain.frx":2FDE0
            Top             =   0
            Width           =   705
         End
         Begin VB.Label lblcurbal 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00472BEA&
            Height          =   195
            Left            =   3720
            TabIndex        =   47
            Top             =   2925
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Course:"
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
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Course Description:"
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
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1410
         End
      End
   End
   Begin VB.Frame framDataViewExport 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9360
      Left            =   3480
      TabIndex        =   92
      Top             =   740
      Width           =   9135
      Begin VB.CheckBox chkSelectAllItems 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select all listed items"
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
         Left            =   120
         TabIndex        =   105
         Top             =   6940
         Width           =   2655
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   9135
         Left            =   120
         ScaleHeight     =   9135
         ScaleWidth      =   8940
         TabIndex        =   93
         Top             =   150
         Width           =   8940
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search Student Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   0
            TabIndex        =   98
            Top             =   0
            Width           =   8895
            Begin VB.PictureBox Picture9 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   1155
               Left            =   40
               ScaleHeight     =   1155
               ScaleWidth      =   8775
               TabIndex        =   99
               Top             =   180
               Width           =   8775
               Begin VB.CommandButton cmdCancelSearchDV 
                  Caption         =   "Cancel Search"
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
                  Left            =   7320
                  MouseIcon       =   "frmMain.frx":31922
                  MousePointer    =   99  'Custom
                  TabIndex        =   103
                  Top             =   430
                  Width           =   1335
               End
               Begin VB.ComboBox cmbSelectCourseExport 
                  Appearance      =   0  'Flat
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
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   100
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.ComboBox cmbSearchCriteria 
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
                  Left            =   2280
                  Style           =   2  'Dropdown List
                  TabIndex        =   101
                  Top             =   480
                  Width           =   2175
               End
               Begin VB.TextBox txtSearchDataView 
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
                  Left            =   4560
                  TabIndex        =   102
                  Top             =   480
                  Width           =   2655
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Search Parameter:"
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
                  Left            =   4560
                  TabIndex        =   173
                  Top             =   120
                  Width           =   1350
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Search with:"
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
                  Left            =   2280
                  TabIndex        =   172
                  Top             =   120
                  Width           =   900
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Course:"
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
                  Left            =   120
                  TabIndex        =   110
                  Top             =   120
                  Width           =   1050
               End
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Data Export Options for Excel (fields to be exported)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   0
            TabIndex        =   94
            Top             =   7155
            Width           =   8895
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   140
               Left            =   4680
               ScaleHeight     =   135
               ScaleWidth      =   4005
               TabIndex        =   158
               Top             =   1200
               Width           =   4000
               Begin VB.Label lblProgress1 
                  BackColor       =   &H00B3F9B8&
                  Height          =   45
                  Left            =   0
                  TabIndex        =   160
                  Top             =   0
                  Width           =   15
               End
               Begin VB.Label lblProgress2 
                  BackColor       =   &H001BD805&
                  Height          =   80
                  Left            =   0
                  TabIndex        =   159
                  Top             =   45
                  Width           =   15
               End
            End
            Begin VB.PictureBox Picture8 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   1575
               Left            =   120
               ScaleHeight     =   1575
               ScaleWidth      =   8700
               TabIndex        =   95
               Top             =   240
               Width           =   8700
               Begin VB.CommandButton cmdFormatDatabaseDV 
                  Caption         =   "Format Database for Data export..."
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
                  MouseIcon       =   "frmMain.frx":31A74
                  MousePointer    =   99  'Custom
                  TabIndex        =   118
                  Top             =   1200
                  Width           =   2895
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Student ID"
                  Enabled         =   0   'False
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
                  Index           =   0
                  Left            =   240
                  TabIndex        =   106
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Exam Index"
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
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   107
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Course"
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
                  Index           =   2
                  Left            =   3360
                  TabIndex        =   108
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "NIC/Passport"
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
                  Index           =   3
                  Left            =   4680
                  TabIndex        =   109
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Name"
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
                  Index           =   4
                  Left            =   6360
                  TabIndex        =   114
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Sex"
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
                  Index           =   5
                  Left            =   240
                  TabIndex        =   111
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Date of Birth"
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
                  Index           =   6
                  Left            =   1680
                  TabIndex        =   112
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Address"
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
                  Index           =   7
                  Left            =   3360
                  TabIndex        =   113
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Contact"
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
                  Index           =   8
                  Left            =   4680
                  TabIndex        =   115
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Email"
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
                  Index           =   9
                  Left            =   6360
                  TabIndex        =   116
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CheckBox chkExportField 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Remarks"
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
                  Index           =   10
                  Left            =   240
                  TabIndex        =   117
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.CommandButton cmdIncluedAll 
                  Caption         =   "Include &All Fields"
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
                  Left            =   5280
                  MouseIcon       =   "frmMain.frx":31BC6
                  MousePointer    =   99  'Custom
                  TabIndex        =   119
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.CommandButton cmdExport 
                  Caption         =   "&Export..."
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
                  Left            =   7600
                  MouseIcon       =   "frmMain.frx":31D18
                  MousePointer    =   99  'Custom
                  TabIndex        =   120
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label lblPro_Indicator 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0%"
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   165
                  Left            =   4200
                  TabIndex        =   162
                  Top             =   940
                  Visible         =   0   'False
                  Width           =   180
               End
               Begin VB.Label lblEx_Pro_Description 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exporting data, please wait..."
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   2460
                  TabIndex        =   161
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   1710
               End
               Begin VB.Image Image4 
                  Height          =   675
                  Left            =   7920
                  Picture         =   "frmMain.frx":31E6A
                  Top             =   120
                  Width           =   660
               End
            End
         End
         Begin MSComctlLib.ListView lvwData 
            Height          =   5055
            Left            =   0
            TabIndex        =   104
            ToolTipText     =   "Double click on a record to edit"
            Top             =   1560
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Student ID"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Exam Index"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Course"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "NIC/Passport"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Name"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sex"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Date of Birth"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Address"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Contact"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Email"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Remarks"
               Object.Width           =   7056
            EndProperty
         End
         Begin MSComctlLib.ImageList imgList 
            Left            =   -240
            Top             =   -360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":335E0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":340DA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":343DB
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Count:"
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
            Left            =   7320
            TabIndex        =   122
            Top             =   6825
            Width           =   1155
         End
         Begin VB.Label lblSelectedCountDV 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   8760
            TabIndex        =   121
            Top             =   6825
            Width           =   105
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00472BEA&
            Height          =   195
            Left            =   3720
            TabIndex        =   96
            Top             =   2925
            Width           =   945
         End
      End
   End
   Begin VB.Frame framEmails 
      BackColor       =   &H00FFFFFF&
      Height          =   9360
      Left            =   3480
      TabIndex        =   62
      Top             =   740
      Width           =   9135
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8955
         Left            =   120
         ScaleHeight     =   8955
         ScaleWidth      =   8940
         TabIndex        =   63
         Top             =   240
         Width           =   8940
         Begin VB.CheckBox chkAllEmail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select all listed items"
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
            Left            =   1800
            TabIndex        =   177
            Top             =   7920
            Width           =   3375
         End
         Begin MSComctlLib.ListView lvwEmail 
            Height          =   7095
            Left            =   1800
            TabIndex        =   176
            Top             =   720
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   12515
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Email Address"
               Object.Width           =   11818
            EndProperty
         End
         Begin VB.CommandButton cmdCopytoClipBoard 
            Caption         =   "Copy to Clipboard"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MouseIcon       =   "frmMain.frx":346DE
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox cmbSelectCourseEmail 
            Appearance      =   0  'Flat
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
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   120
            Width           =   4455
         End
         Begin VB.CommandButton cmdSavetofile 
            Caption         =   "Save to file"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MouseIcon       =   "frmMain.frx":34830
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdMailClient 
            Caption         =   "Mail Client"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":34982
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   4920
            Width           =   1455
         End
         Begin VB.Label lblCountEmail 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   8670
            TabIndex        =   77
            Top             =   7920
            Width           =   105
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Count:"
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
            Left            =   7200
            TabIndex        =   76
            Top             =   7920
            Width           =   1155
         End
         Begin VB.Image Image6 
            Height          =   690
            Left            =   0
            Picture         =   "frmMain.frx":34AD4
            Top             =   7440
            Width           =   735
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00472BEA&
            Height          =   195
            Left            =   3720
            TabIndex        =   65
            Top             =   2925
            Width           =   945
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course:"
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
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   1050
         End
      End
   End
   Begin VB.Frame framContacts 
      BackColor       =   &H00FFFFFF&
      Height          =   9360
      Left            =   3480
      TabIndex        =   58
      Top             =   740
      Width           =   9135
      Begin VB.CheckBox chkAllContact 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select all listed items"
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
         Left            =   1920
         TabIndex        =   175
         Top             =   8160
         Width           =   3375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8955
         Left            =   120
         ScaleHeight     =   8955
         ScaleWidth      =   8940
         TabIndex        =   59
         Top             =   240
         Width           =   8940
         Begin VB.PictureBox Picture14 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   140
            Left            =   4800
            ScaleHeight     =   135
            ScaleWidth      =   4005
            TabIndex        =   178
            Top             =   8640
            Width           =   4000
            Begin VB.Label lblProgress2_Contact 
               BackColor       =   &H001BD805&
               Height          =   80
               Left            =   0
               TabIndex        =   180
               Top             =   45
               Width           =   15
            End
            Begin VB.Label lblProgress1_Contact 
               BackColor       =   &H00B3F9B8&
               Height          =   45
               Left            =   0
               TabIndex        =   179
               Top             =   0
               Width           =   15
            End
         End
         Begin MSComctlLib.ListView lvwName_Contact 
            Height          =   7095
            Left            =   1800
            TabIndex        =   174
            Top             =   720
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   12515
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   6526
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Contact Number/s"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.CommandButton cmdSaveName 
            Caption         =   "Save Names"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":365AE
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox cmbSelectCourseContact 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   120
            Width           =   4435
         End
         Begin VB.CommandButton cmdSaveNameContact 
            Caption         =   "Save Contacts"
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
            Left            =   120
            MouseIcon       =   "frmMain.frx":36700
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblProgress_Indicator_Contact 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Left            =   4410
            TabIndex        =   182
            Top             =   8640
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblProgress_Des_Contact 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saving data, please wait..."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2805
            TabIndex        =   181
            Top             =   8640
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Count:"
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
            Left            =   7200
            TabIndex        =   79
            Top             =   7920
            Width           =   1155
         End
         Begin VB.Label lblCountName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   8660
            TabIndex        =   78
            Top             =   7920
            Width           =   105
         End
         Begin VB.Image Image7 
            Height          =   780
            Left            =   0
            Picture         =   "frmMain.frx":36852
            Top             =   7360
            Width           =   630
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Course:"
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
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00472BEA&
            Height          =   195
            Left            =   3720
            TabIndex        =   60
            Top             =   2925
            Width           =   945
         End
      End
   End
   Begin VB.Menu mnuSavetoClipBoard 
      Caption         =   "Save_ClipBoard"
      Visible         =   0   'False
      Begin VB.Menu mnuFormattedtoEmailClipboard 
         Caption         =   "&Formatted to Email..."
      End
      Begin VB.Menu mnuEmailClipboardNormal 
         Caption         =   "&Without Format..."
      End
   End
   Begin VB.Menu mnuSavetoFile 
      Caption         =   "Save_File"
      Visible         =   0   'False
      Begin VB.Menu mnuFormattedtoEmailFile 
         Caption         =   "F&ormatted to Email..."
      End
      Begin VB.Menu mnuEmailFileNormal 
         Caption         =   "W&ithout Format..."
      End
   End
   Begin VB.Menu mnuSavetoTextExcel 
      Caption         =   "SavetoTextExcel"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveastextfile 
         Caption         =   "Save as Text File..."
      End
      Begin VB.Menu mnuSaveasmsexcelfile 
         Caption         =   "Save as Ms Excel File..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================
'=  Developed by P. G. B. Prasanna       =
'=  A Software Developer from Sri Lanka  =
'=  E-Mail: pgbsoft@gmail.com            =
'=========================================

Dim rststdinfo As ADODB.Recordset: Dim rststdinfosearch As ADODB.Recordset
Dim rstcourseinfo As ADODB.Recordset: Dim rststdinfodv As ADODB.Recordset

Dim blnadd As Boolean: Dim blnedit As Boolean
Dim blnstdadd As Boolean: Dim blnstdedit As Boolean
Dim blnstdSearch As Boolean

Dim firststd, secondstd, firstcourse, secondcourse As String
Dim f_name, filesavepath As String
Dim Find_val As String
Dim strAppend As String


Dim selectedemail, selectedname, Load_State As Integer
Dim intAdvanced_search, intlogoff As Integer
Dim intDbClick As Integer

Dim lngProgressLimit As Long
Dim intProgressChange As Integer
Dim l As Integer
Dim File_Num As Long


Private Sub chkAllContact_Click()
If chkAllContact.Value = 1 Then
    For i = 1 To lvwName_Contact.ListItems.count
    lvwName_Contact.ListItems(i).Checked = True
    Check_List_View_Selected lvwName_Contact, 1
    Next
Else
    For i = 1 To lvwName_Contact.ListItems.count
    lvwName_Contact.ListItems(i).Checked = False
    lblCountName.Caption = "0"
    Next
End If

End Sub

Private Sub chkAllContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub chkAllEmail_Click()
If chkAllEmail.Value = 1 Then
    For i = 1 To lvwEmail.ListItems.count
    lvwEmail.ListItems(i).Checked = True
    Check_List_View_Selected lvwEmail, 2
    Next
Else
    For i = 1 To lvwEmail.ListItems.count
    lvwEmail.ListItems(i).Checked = False
    lblCountEmail.Caption = "0"
    Next
End If

End Sub

Private Sub chkAllEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub chkOpenFile_Click()
On Error Resume Next
Call Write_Registry(File_Open_Set, chkOpenFile.Value, 1)
strFile_Open = chkOpenFile.Value
End Sub
Private Sub cmbCourse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtNICPassport.SetFocus
End Sub
Private Sub cmbSearchCriteria_Click()
cmbSelectCourseExport_Click
End Sub

Private Sub cmbSelectCourseContact_Click()
On Error Resume Next
Dim rstselectcourse_contact As ADODB.Recordset

lvwName_Contact.ListItems.Clear
lvwName_Contact.SmallIcons = imgList
chkAllContact.Value = 0: lblCountName.Caption = 0

Dim strselectcourse_contact As String

If cmbSelectCourseContact = "---All Courses---" Then
    Set rstselectcourse_contact = New ADODB.Recordset
            rstselectcourse_contact.Open "SELECT* FROM STUDENTINFO ORDER BY STD_COURSE", dbcon, adOpenStatic, adLockOptimistic
            If rstselectcourse_contact.RecordCount > 0 Then
               Do While Not rstselectcourse_contact.EOF
                    Set i = lvwName_Contact.ListItems.Add(, , rstselectcourse_contact("STD_NAME"))
                        i.SubItems(1) = rstselectcourse_contact("STD_CONTACT")
                        i.SmallIcon = 3
                   
               'DoEvents
               rstselectcourse_contact.MoveNext
               Loop
            End If
Else
    strselectcourse_contact = cmbSelectCourseContact
            Set rstselectcourse_contact = New ADODB.Recordset
            rstselectcourse_contact.Open "SELECT* FROM STUDENTINFO WHERE[STD_COURSE] = '" & strselectcourse_contact & "' ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
            If rstselectcourse_contact.RecordCount > 0 Then
               Do While Not rstselectcourse_contact.EOF
                    Set i = lvwName_Contact.ListItems.Add(, , rstselectcourse_contact("STD_NAME"))
                        i.SubItems(1) = rstselectcourse_contact("STD_CONTACT")
                        i.SmallIcon = 3
                   
               'DoEvents
               rstselectcourse_contact.MoveNext
               Loop
            End If
rstselectcourse_contact.Close
Set rstselectcourse_contact = Nothing
End If
End Sub

Private Sub cmbSelectCourseEmail_Click()
On Error Resume Next
Dim rstselectcourse_email As ADODB.Recordset
Dim strselectcourse_email, valid_email As String
lstEmail.Clear

lvwEmail.ListItems.Clear
lvwEmail.SmallIcons = imgList
chkAllEmail.Value = 0: lblCountEmail.Caption = 0

If cmbSelectCourseEmail = "---All Courses---" Then
    Set rstselectcourse_email = New ADODB.Recordset
            rstselectcourse_email.Open "SELECT* FROM STUDENTINFO ORDER BY STD_COURSE", dbcon, adOpenStatic, adLockOptimistic
            If rstselectcourse_email.RecordCount > 0 Then
               Do While Not rstselectcourse_email.EOF
                   valid_email = rstselectcourse_email("STD_EMAIL")
                       If InStr(1, valid_email, "@", vbTextCompare) Then
                        'lstEmail.AddItem rstselectcourse_email("STD_EMAIL")
                           'lstEmail.AddItem valid_email
                            Set i = lvwEmail.ListItems.Add(, , valid_email)
                            i.SmallIcon = 2
                       End If
               rstselectcourse_email.MoveNext
               Loop
            End If
Else
    strselectcourse_email = cmbSelectCourseEmail
            Set rstselectcourse_email = New ADODB.Recordset
            rstselectcourse_email.Open "SELECT* FROM STUDENTINFO WHERE[STD_COURSE] = '" & strselectcourse_email & "' ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
            If rstselectcourse_email.RecordCount > 0 Then
               Do While Not rstselectcourse_email.EOF
                   valid_email = rstselectcourse_email("STD_EMAIL")
                       If InStr(1, valid_email, "@", vbTextCompare) Then
                           'lstEmail.AddItem rstselectcourse_email("STD_EMAIL")
                           Set i = lvwEmail.ListItems.Add(, , valid_email)
                           i.SmallIcon = 2
                       End If
               rstselectcourse_email.MoveNext
               Loop
            End If
rstselectcourse_email.Close
Set rstselectcourse_email = Nothing
End If
End Sub

Private Sub cmbSex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtDateofBirth.SetFocus
End Sub

Private Sub cmdCancelDV_Click()
'txtSearchDataView = ""
cmbSelectCourseExport = cmbSelectCourseExport.List(0)
End Sub

Private Sub cmdCancelSearch_Click()
cmdCancelSearch.Enabled = False
txtSearch = ""
Open_Student_Info
End Sub

Private Sub cmdCancelSearchDV_Click()
txtSearchDataView = ""
cmbSelectCourseExport = cmbSelectCourseExport.List(0)
End Sub

Private Sub cmdClear_Click()
Dim rstclearcourseinfo As ADODB.Recordset
If User <> "Administrator" Then
    MsgBox "Only the User: Administrator can perform this action." & vbCrLf & "Please, contact Administrator.", vbExclamation
    Exit Sub
End If
Dim lngProcessing_Record As Long
If cmbClearDatabase = cmbClearDatabase.List(0) Then
    If MsgBox("Are you sure you want to clear " & cmbClearDatabase.List(0) & "?", vbInformation + vbYesNo) = vbYes Then
        On Error Resume Next
        lngProcessing_Record = 0
        Reset_Progress 7
        Control_Enable_With_Progress False
        
        Set rstclearcourseinfo = New ADODB.Recordset
        rstclearcourseinfo.CursorLocation = adUseClient
        rstclearcourseinfo.Open "SELECT * FROM COURSEINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
            If rstclearcourseinfo.RecordCount > 0 Then
                lngProgressLimit = rstclearcourseinfo.RecordCount
                Do While Not rstclearcourseinfo.EOF
                    lngProcessing_Record = lngProcessing_Record + 1
                    intProgressChange = (lngProcessing_Record / lngProgressLimit) * 100
                    rstclearcourseinfo.Delete
                    rstclearcourseinfo.MoveNext
                    
                    lblProgressImport1.Width = (intProgressChange / 100) * 4000
                    lblProgressImport2.Width = (intProgressChange / 100) * 4000
                    lblImportIndicator = intProgressChange & "%"
                    DoEvents
                Loop
            End If
        If lblImportIndicator.Caption = 100 & "%" Then: lblImportDescription.Caption = "Data was deleted..."
        Control_Enable_With_Progress True
        MsgBox "Selected information cleared.", vbInformation
        Reset_Progress 5
        rstclearcourseinfo.Close
        Set rstclearcourseinfo = Nothing
    End If
ElseIf cmbClearDatabase = cmbClearDatabase.List(1) Then
    If MsgBox("Are you sure you want to clear " & cmbClearDatabase.List(1) & "?", vbInformation + vbYesNo) = vbYes Then
        On Error Resume Next
        lngProcessing_Record = 0
        Reset_Progress 7
        Control_Enable_With_Progress False
        Set rstclearstdinfo = New ADODB.Recordset
        rstclearstdinfo.CursorLocation = adUseClient
        rstclearstdinfo.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
            If rstclearstdinfo.RecordCount > 0 Then
                lngProgressLimit = rstclearstdinfo.RecordCount
                Do While Not rstclearstdinfo.EOF
                    lngProcessing_Record = lngProcessing_Record + 1
                    intProgressChange = (lngProcessing_Record / lngProgressLimit) * 100
                    rstclearstdinfo.Delete
                    rstclearstdinfo.MoveNext
                    
                    lblProgressImport1.Width = (intProgressChange / 100) * 4000
                    lblProgressImport2.Width = (intProgressChange / 100) * 4000
                    lblImportIndicator = intProgressChange & "%"
                    DoEvents
                Loop
            End If
        If lblImportIndicator.Caption = 100 & "%" Then: lblImportDescription.Caption = "Data was deleted..."
        Control_Enable_With_Progress True
        MsgBox "Selected information cleared.", vbInformation
        Reset_Progress 5
        rstclearstdinfo.Close
        Set rstclearstdinfo = Nothing
    End If
End If
End Sub

Private Sub cmdCopytoClipBoard_Click()
PopupMenu mnuSavetoClipBoard, 50, 3700, 4400, mnuFormattedtoEmailClipboard
End Sub

Private Sub cmdDefaultFLocation_Click()
frmBrowseFolder.Show 1
End Sub

Private Sub cmdFormatDatabase_Click()

'The following process calls the function 'Format_Data_Field' to remove all the unnecessary characters like new line characters entered
'in some fields like student name, Address, Remarks.

Dim rstFormat_Stdinfo As ADODB.Recordset
Dim lngProcessing_Record As Long

If MsgBox("Are you sure you need to format the database for exporting data?", vbQuestion + vbYesNo) = vbYes Then
    On Error Resume Next
    lngProcessing_Record = 0
    Reset_Progress 6
    Control_Enable_With_Progress False
    
    Set rstFormat_Stdinfo = New ADODB.Recordset
    rstFormat_Stdinfo.CursorLocation = adUseClient
    rstFormat_Stdinfo.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
    
    'get the progress Limit
    lngProgressLimit = rstFormat_Stdinfo.RecordCount
        Do While Not rstFormat_Stdinfo.EOF
            'Progress Value Handling
            lngProcessing_Record = lngProcessing_Record + 1
            intProgressChange = (lngProcessing_Record / lngProgressLimit) * 100
               
            'processing Student Name field
            rstFormat_Stdinfo("STD_NAME") = Format_Data_Field(Trim(rstFormat_Stdinfo("STD_NAME")))
            'processing Student Address field
            rstFormat_Stdinfo("STD_ADDRESS") = Format_Data_Field(Trim(rstFormat_Stdinfo("STD_ADDRESS")))
            'processing Remarks field
            rstFormat_Stdinfo("REMARKS") = Format_Data_Field(Trim(rstFormat_Stdinfo("REMARKS")))
               
            rstFormat_Stdinfo.Update
            rstFormat_Stdinfo.MoveNext
           
            lblProgress1.Width = (intProgressChange / 100) * 4000
            lblProgress2.Width = (intProgressChange / 100) * 4000
            lblProgressImport1.Width = (intProgressChange / 100) * 4000
            lblProgressImport2.Width = (intProgressChange / 100) * 4000
            lblPro_Indicator = intProgressChange & "%"
            lblImportIndicator = intProgressChange & "%"
            DoEvents
        Loop
    If lblPro_Indicator.Caption = 100 & "%" Then: lblEx_Pro_Description.Caption = "Formatting data completed..."
    If lblImportIndicator.Caption = 100 & "%" Then: lblImportDescription.Caption = "Formatting data completed..."
    Control_Enable_With_Progress True
    MsgBox "Database format completed.", vbInformation
    Reset_Progress 5
    rstFormat_Stdinfo.Close
    Set rstFormat_Stdinfo = Nothing
End If
End Sub

Private Sub cmdImportData_Click()
If User <> "Administrator" Then
    MsgBox "Only the User: Administrator can perform this action." & vbCrLf & "Please, contact Administrator.", vbExclamation
    Exit Sub
End If

On Error GoTo Err
Dim rstMain As ADODB.Recordset
Dim rstImport As ADODB.Recordset
Dim rstCheck As ADODB.Recordset

Dim r_count As Long
Dim lngValid_count As Long
Dim lngInvalid_count As Long
Dim rec_val As String

If txtImportDblocation = "" Then
    MsgBox "Please select the database...", vbExclamation
    cmdSelectDatabase_Click
    Exit Sub
End If

openImportDB

If intproceed = 0 Then: Exit Sub

r_count = 0
lngValid_count = 0

'lngInvalid_count = 0

Set rstMain = New ADODB.Recordset
    rstMain.CursorLocation = adUseClient
    rstMain.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
Set rstImport = New ADODB.Recordset
    rstImport.CursorLocation = adUseClient
    rstImport.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbimport, adOpenStatic, adLockOptimistic
    
lngProgressLimit = rstImport.RecordCount
Reset_Progress 4
Control_Enable_With_Progress False

Do While Not rstImport.EOF
    r_count = r_count + 1
    intProgressChange = (r_count / lngProgressLimit) * 100
    
    rec_val = rstImport("STD_ID")
    If Check_for_Record_Existence("STUDENTINFO", "STD_ID", rec_val, 1) = False Then
        lngValid_count = lngValid_count + 1
        On Error Resume Next
        rstMain.AddNew
        rstMain("STD_ID") = rstImport("STD_ID")
        rstMain("STD_EXAM_INDEX") = rstImport("STD_EXAM_INDEX")
        rstMain("STD_COURSE") = rstImport("STD_COURSE")
        rstMain("STD_NIC_PASSPORT") = rstImport("STD_NIC_PASSPORT")
        rstMain("STD_NAME") = Format_Data_Field(rstImport("STD_NAME"))
        rstMain("SEX") = rstImport("SEX")
        rstMain("DATE_OF_BIRTH") = rstImport("DATE_OF_BIRTH")
        rstMain("STD_ADDRESS") = Format_Data_Field(rstImport("STD_ADDRESS"))
        rstMain("STD_CONTACT") = rstImport("STD_CONTACT")
        rstMain("STD_EMAIL") = rstImport("STD_EMAIL")
        rstMain("REMARKS") = Format_Data_Field(rstImport("REMARKS"))
        rstMain.Update
    End If
        rstImport.MoveNext
        lblProgressImport1.Width = (intProgressChange / 100) * 4000
        lblProgressImport2.Width = (intProgressChange / 100) * 4000
        lblImportIndicator.Caption = intProgressChange & "%"
        DoEvents
Loop
If lblImportIndicator.Caption = 100 & "%" Then: lblImportDescription.Caption = "Data import completed..."

rstMain.Close
Set rstMain = Nothing
rstImport.Close
Set rstImport = Nothing

lngInvalid_count = r_count - lngValid_count

Control_Enable_With_Progress True

MsgBox "Data Import was done." & vbCrLf & vbCrLf & _
       "Data Import Status:" & vbCrLf & _
       "------------------------" & vbCrLf & _
       r_count & " records found." & vbCrLf & _
       lngValid_count & " records imported." & vbCrLf & _
       lngInvalid_count & " ignored as duplicate.", vbInformation
Reset_Progress 3

dbimport.Close
Set dbimport = Nothing
Exit Sub
Err:
MsgBox "This database is not supported...", vbExclamation
'lblImportIndicator.Visible = False: lblImportDescription.Visible = False
Reset_Progress 3
End Sub

Private Sub cmdMailClient_Click()
On Error GoTo Err
Open_File_Or_Location (strMail_Client)
Exit Sub
Err:
MsgBox "Please set the Mail Client Application.", vbCritical
frmBF.Show 1
End Sub

Private Sub cmdSaveName_Click()
If Check_List_View_Selected(lvwName_Contact) = False Then: MsgBox "Please select items from the list.", vbExclamation: Exit Sub

File_Num = FreeFile

On Error Resume Next
If Dir(File_Root, vbDirectory) <> "" Then
    filesavepath = File_Root & "Names-" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM")
    Open filesavepath & ".txt" For Output As #File_Num
        Print #File_Num, File_Clipboard_Handle_Return_String(lvwName_Contact)
    Close #File_Num
    MsgBox "File successfully saved to " & filesavepath & ".", vbInformation
    If strFile_Open = 1 Then: Open_File_Or_Location (filesavepath)

Else
    MsgBox "Default file saving location not found...", vbCritical
    frmBrowseFolder.Show 1
    Exit Sub
End If
End Sub

Private Sub cmdSaveNameContact_Click()
PopupMenu mnuSavetoTextExcel, 30, 6100, 4930, Me.mnuSaveastextfile
End Sub

Private Sub cmdSavetofile_Click()
PopupMenu mnuSavetoFile, 50, 3700, 5720, mnuFormattedtoEmailFile
End Sub


Private Sub cmdSelectDatabase_Click()
intbrowseoption = 1
frmBF.Show 1
End Sub

Private Sub cmdSetasdefault_Click()
On Error Resume Next
    Call Write_Registry(Default_Prompt, cmbLoadPrompt.ListIndex, 1)
    strEmail_Seperator = Left(cmbEmailSeperator, 1)
End Sub

Private Sub cmdSetasDefaultEmSeperator_Click()
On Error Resume Next
    strEmail_Seperator = Left(cmbEmailSeperator, 1)
    Call Write_Registry(Email_Seperator, cmbEmailSeperator)
End Sub

Private Sub cmdStdAdd_Click()
'Enable_Controls
'On Error Resume Next
If Privilege_Proceed = False Then: Exit Sub
    blnstdadd = True
    'cmdStdSearch.Enabled = False
    'cmdCancelSearch.Enabled = False
    txtSearch.Text = ""
    Clear_Fields
    Button_Add_Edit_Save_Cancle_RecordExist_Mode False
    txtStdID.SetFocus
End Sub

Private Sub cmdStdCancel_Click()
If MsgBox("Are you sure you need to cancel?", vbYesNo + vbQuestion) = vbYes Then
    On Error Resume Next
    rststdinfo.CancelUpdate
    Clear_Fields
    Inforfield
    Disable_Controls
    blnstdadd = False
    blnstdedit = False
End If
End Sub

Private Sub cmdStdDelete_Click()
'blndelete_Click = True
If Privilege_Proceed = False Then: Exit Sub

If MsgBox("Delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
    On Error Resume Next
    If intSearch_delete = 1 Then: rststdinfo.Delete: cmdCancelSearch_Click: Exit Sub
       rststdinfo.Delete
       rststdinfo.MoveNext
           If rststdinfo.EOF Then
               rststdinfo.MoveLast
           End If
       Inforfield
End If
End Sub

Private Sub cmdStdEdit_Click()
If Privilege_Proceed = False Then: Exit Sub
    blnstdedit = True
    Enable_Controls
    txtSearch.Text = ""
    txtStdID.SetFocus
    cmdSearch.Enabled = False
    firststd = txtStdID
    Button_Add_Edit_Save_Cancle_RecordExist_Mode False
End Sub

Private Sub cmdStdFirst_Click()
   If rststdinfo.BOF = False Then
        rststdinfo.MoveFirst
        Inforfield
        MsgBox "You are on the First Record.", vbInformation
   End If
End Sub

Private Sub cmdStdLast_Click()
  If rststdinfo.EOF = False Then
        rststdinfo.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
  End If
End Sub

Private Sub cmdStdNext_Click()
On Error Resume Next
   If rststdinfo.EOF = False Then
        rststdinfo.MoveNext
        Inforfield
   ElseIf rststdinfo.EOF Then
        rststdinfo.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
   End If
End Sub

Private Sub cmdStdPrevious_Click()
   If rststdinfo.BOF = False Then
        rststdinfo.MovePrevious
        Inforfield
   ElseIf rststdinfo.BOF = True Then
        rststdinfo.MoveFirst
        Inforfield
        MsgBox "You are on the first Record.", vbInformation
   End If
End Sub
Private Sub cmdSearch_Click()
On Error GoTo Err
'rststdinfo.Close
'Set rststdinfo = Nothing
If txtSearch.Text = "" Then
    MsgBox "Search field is empty...!", vbExclamation
    txtSearch.SetFocus
    Exit Sub
End If
blnstdSearch = True
'Set rststdinfosearch = New ADODB.Recordset
    'rststdinfosearch.CursorLocation = adUseClient
    'rststdinfosearch.Open "SELECT * FROM CusInfo ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
Dim Find_val As String
Find_val = txtSearch
'Set rststdinfo = New ADODB.Recordset
'rststdinfo.Open "SELECT* FROM STUDENTINFO WHERE[STD_ID] = '" & Find_val & "'", dbcon, adOpenStatic, adLockOptimistic
    intSearch_delete = 1
    Select Case cmbSearchOption.ListIndex
        Case 0: Set rststdinfo = New ADODB.Recordset
                rststdinfo.Open "SELECT* FROM STUDENTINFO WHERE[STD_ID] = '" & Find_val & "'", dbcon, adOpenStatic, adLockOptimistic: Inforfield
        Case 1: Set rststdinfo = New ADODB.Recordset
                rststdinfo.Open "SELECT* FROM STUDENTINFO WHERE[STD_EXAM_INDEX] = '" & Find_val & "'", dbcon, adOpenStatic, adLockOptimistic: Inforfield
        Case 2: Set rststdinfo = New ADODB.Recordset
                rststdinfo.Open "SELECT* FROM STUDENTINFO WHERE[STD_NIC_PASSPORT] = '" & Find_val & "'", dbcon, adOpenStatic, adLockOptimistic: Inforfield
        Case 3: Set rststdinfo = New ADODB.Recordset
                rststdinfo.Open "SELECT* FROM STUDENTINFO WHERE[STD_NAME] Like '" & Find_val & "%'", dbcon, adOpenStatic, adLockOptimistic: Inforfield
    End Select
Inforfield
Exit Sub
Err:
MsgBox "Error occured: " & Err.Description
End Sub

Private Sub cmdStdSave_Click()
Dim rstRecordExist As ADODB.Recordset
Dim Find_recordexist As String
Const MsgBlank = "Field can not be blank...!"
On Error Resume Next
If blnstdadd = True Then
    If txtStdID = "" Then: MsgBox MsgBlank, vbExclamation: txtStdID.SetFocus: Exit Sub
    If cmbCourse = "" Then: MsgBox MsgBlank, vbExclamation: cmbCourse.SetFocus: Exit Sub
    If txtStdName = "" Or InStr(1, txtStdName, Chr(10)) = 2 Then: MsgBox MsgBlank, vbExclamation: txtStdName = "": txtStdName.SetFocus: Exit Sub
    If txtStdAddress = "" Or InStr(1, txtStdAddress, Chr(10)) = 2 Then: MsgBox MsgBlank, vbExclamation: txtStdAddress = "": txtStdAddress.SetFocus: Exit Sub
    If txtStdTel = "" Then: MsgBox MsgBlank, vbExclamation: txtStdTel.SetFocus: Exit Sub
    If txtStdEmail = "" Then: MsgBox MsgBlank, vbExclamation: txtStdEmail.SetFocus: Exit Sub
    If txtExamIndex = "" Then: MsgBox MsgBlank, vbExclamation: txtExamIndex.SetFocus: Exit Sub
    If txtNICPassport = "" Then: MsgBox MsgBlank, vbExclamation: txtNICPassport.SetFocus: Exit Sub
    If cmbSex = "" Then: MsgBox MsgBlank, vbExclamation: cmbSex.SetFocus: Exit Sub
    If txtStdAddress = "" Then: MsgBox MsgBlank, vbExclamation: txtStdAddress.SetFocus: Exit Sub
    
    Find_recordexist = txtStdID
    
    If Check_for_Record_Existence("STUDENTINFO", "STD_ID", Find_recordexist) = True Then: txtStdID.SelStart = 0: txtStdID.SelLength = Len(txtStdID): txtStdID.SetFocus: Exit Sub
    
    rststdinfo.AddNew
    Save_Data
            
ElseIf blnstdedit = True Then
    If txtStdID = "" Then: MsgBox MsgBlank, vbExclamation: txtStdID.SetFocus: Exit Sub
    If txtStdEmail = "" Then: MsgBox MsgBlank, vbExclamation: txtStdEmail.SetFocus: Exit Sub
    secondstd = txtStdID
        If firststd <> secondstd Then
            Find_recordexist = txtStdID
            If Check_for_Record_Existence("STUDENTINFO", "STD_ID", Find_recordexist) = True Then: txtStdID.SelStart = 0: txtStdID.SelLength = Len(txtStdID): txtStdID.SetFocus: Exit Sub
        End If
    Save_Data
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
txtSearch.SetFocus
intbrowseoption = 0
End Sub

Private Sub Form_Load()
Me.Picture = frmBG_Image.imgBG_Image.Picture

Dim intFirstStart As Integer
'Load_Resource_Picture_Enable True
Me.Top = ((Screen.Height - Me.Height) / 2) - 150
Me.Left = (Screen.Width - Me.Width) / 2
lblCurrentUser = User
intlogoff = 0

Load_Course_Info
Load_Options_Info
intbrowseoption = 0

On Error Resume Next
strMail_Client = Read_Registry(Email_Client_Loc)
strFile_Save = Read_Registry(Default_File_Loc)
strFile_Open = Read_Registry(File_Open_Set)
Load_State = Read_Registry(Default_Prompt)
intFirstStart = Val(Read_Registry(Is_First))
strEmail_Seperator = Left(Read_Registry(Email_Seperator), 1)

If strFile_Open = "" And strFile_Open <> "1" And strFile_Open <> "0" Then: strFile_Open = "1": chkOpenFile.Value = 1
If intFirstStart = 0 Then: Call Write_Registry(Is_First, 1, 1): Load_State = 5: cmbLoadPrompt = cmbLoadPrompt.List(5)
If strEmail_Seperator = "" And strEmail_Seperator <> "," And strEmail_Seperator <> ";" Then: strEmail_Seperator = ","

Select Case Load_State
    Case 0: lblViewExport_MouseUp 0, 0, 0, 0
    Case 1: lblGetEmails_MouseUp 0, 0, 0, 0
    Case 2: lblGetContacts_MouseUp 0, 0, 0, 0
    Case 3: lblStudentInfo_MouseUp 0, 0, 0, 0
    Case 4: lblCourseInfo_MouseUp 0, 0, 0, 0
    Case Else: lblOptions_MouseUp 0, 0, 0, 0
End Select
Check_Privilege_Commands
Unload frmSplash
End Sub
Public Sub Inforfield()
On Error Resume Next
cmdSearch.Enabled = True
 If rststdinfo.RecordCount > 0 Then
    txtStdID = rststdinfo("STD_ID")
    txtExamIndex = rststdinfo("STD_EXAM_INDEX")
    cmbCourse = rststdinfo("STD_COURSE")
    txtNICPassport = rststdinfo("STD_NIC_PASSPORT")
    txtStdName = rststdinfo("STD_NAME")
    cmbSex = rststdinfo("SEX")
    txtDateofBirth = rststdinfo("DATE_OF_BIRTH")
    txtStdAddress = rststdinfo("STD_ADDRESS")
    txtStdTel = rststdinfo("STD_CONTACT")
    txtStdEmail = rststdinfo("STD_EMAIL")
    txtRemarks = rststdinfo("REMARKS")
    
    Button_Add_Edit_Save_Cancle_RecordExist_Mode True
    If blnstdSearch = True Then
        cmdCancelSearch.Enabled = True
    End If
 Else
       If blnstdSearch = True Then
            blnstdSearch = False
            MsgBox "Record not found...", vbExclamation
            lblStudentInfo_MouseUp 0, 0, 0, 0
            txtSearch.SetFocus
            Exit Sub
       End If
       Clear_Fields
       Disable_Controls
       Button_Record_Not_Exist_Mode
       dtpDate.Value = Date
       
       Exit Sub
 End If
 
On Error Resume Next
txtSearch.SetFocus
If blnstdSearch = True Then
    blnstdSearch = False
End If
'rststdinfo.Close
'Set rststdinfo = Nothing
End Sub

Sub Save_Data()
On Error GoTo Err
   rststdinfo("STD_ID") = txtStdID
   rststdinfo("STD_EXAM_INDEX") = UCase(Trim(txtExamIndex))
   rststdinfo("STD_COURSE") = UCase(Trim(cmbCourse))
   rststdinfo("STD_NIC_PASSPORT") = UCase(Trim(txtNICPassport))
   'formatting Student Name field so that it can not contain unnessasary characters.
   rststdinfo("STD_NAME") = Format_Data_Field(Trim(txtStdName))
   rststdinfo("SEX") = cmbSex
   rststdinfo("DATE_OF_BIRTH") = Trim(txtDateofBirth)
   'formatting Student Address field so that it can not contain unnessasary characters.
   rststdinfo("STD_ADDRESS") = Format_Data_Field(Trim(txtStdAddress))
   rststdinfo("STD_CONTACT") = Trim(txtStdTel)
   rststdinfo("STD_EMAIL") = Trim(txtStdEmail)
   'formatting Remarks field so that it can not contain unnessasary characters.
   rststdinfo("REMARKS") = Format_Data_Field(Trim(txtRemarks))
   rststdinfo.Update
   Button_Add_Edit_Save_Cancle_RecordExist_Mode True
   Disable_Controls
   blnstdadd = False
   blnstdedit = False
Exit Sub
Err:
MsgBox Err.Description, vbCritical
End Sub

Public Sub Button_Add_Edit_Save_Cancle_RecordExist_Mode(bval As Boolean)
cmdStdCancel.Enabled = Not bval: cmdStdSave.Enabled = Not bval
cmdStdAdd.Enabled = bval: cmdStdEdit.Enabled = bval
cmdStdDelete.Enabled = bval: cmdStdPrevious.Enabled = bval
cmdStdNext.Enabled = bval: cmdStdFirst.Enabled = bval: cmdStdLast.Enabled = bval
End Sub

Public Sub Button_Record_Not_Exist_Mode()
cmdStdAdd.Enabled = True: cmdStdEdit.Enabled = False
cmdStdDelete.Enabled = False: cmdStdCancel.Enabled = False
cmdStdSave.Enabled = False: cmdStdPrevious.Enabled = False
cmdStdNext.Enabled = False: cmdStdFirst.Enabled = False: cmdStdLast.Enabled = False
End Sub

Public Sub Clear_Fields()
Enable_Controls
  txtStdID = "":  txtStdName = ""
  txtStdAddress = "": txtStdTel = ""
  txtStdEmail = "": txtRemarks = ""
  txtExamIndex = "": txtNICPassport = ""
  txtDateofBirth = ""
End Sub
Public Sub Disable_Controls()
 txtStdID.Enabled = False: cmbCourse.Enabled = False
 txtStdName.Enabled = False: txtStdAddress.Enabled = False
 txtStdTel.Enabled = False: txtStdEmail.Enabled = False
 txtRemarks.Enabled = False: txtExamIndex.Enabled = False
 txtNICPassport.Enabled = False: txtDateofBirth.Enabled = False
 cmbSex.Enabled = False
End Sub

Public Sub Enable_Controls()
 txtStdID.Enabled = True: cmbCourse.Enabled = True
 txtStdName.Enabled = True: txtStdAddress.Enabled = True
 txtStdTel.Enabled = True: txtStdEmail.Enabled = True
 txtRemarks.Enabled = True: txtExamIndex.Enabled = True
 txtNICPassport.Enabled = True: txtDateofBirth.Enabled = True
 cmbSex.Enabled = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub
Private Sub framContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub framCourseInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub


Private Sub framDataViewExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub framEmails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub framOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub framStudentInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub cmdFormatDatabaseDV_Click()
cmdFormatDatabase_Click
cmbSelectCourseExport = cmbSelectCourseExport.List(0)
End Sub

Private Sub cmdFormatDatabaseDV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit_MouseUp 0, 0, 0, 0
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblViewExport_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblViewExport_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblViewExport_MouseUp 0, 0, 0, 0
End Sub
Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetContacts_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetContacts_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetContacts_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCreateUser_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCreateUser_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblManageUser_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblManageUser_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStudentInfo_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStudentInfo_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStudentInfo_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetEmails_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetEmails_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetEmails_MouseUp 0, 0, 0, 0
End Sub


Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCourseInfo_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCourseInfo_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCourseInfo_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackupImport_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackupImport_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub
Private Sub Image20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Image21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogOff_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogOff_MouseUp 0, 0, 0, 0
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub
Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOptions_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOptions_MouseMove 1, 1, 1, 1
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOptions_MouseUp 0, 0, 0, 0
End Sub

Private Sub Label12_Click()
Mail_Me "pgbsoft@gmail.com"
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label12.ForeColor = &H6C76F2 Then
    Label12.ForeColor = vbBlue
Else: Label12.ForeColor = &H6C76F2
End If
End Sub


Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus

End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label48.Top = Label48.Top + 20
Label48.Left = Label48.Left + 20
End Sub

Private Sub Label48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label48.ForeColor = vbRed Then
    Label48.ForeColor = vbBlue
ElseIf Label48.ForeColor = vbBlue Then
       Label48.ForeColor = vbRed
End If
End Sub

Private Sub Label48_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label48.Top = Label48.Top - 20
Label48.Left = Label48.Left - 20

If MsgBox("Are you sure you need to clear the Database Iocation ?", vbYesNo + vbQuestion) = vbYes Then
    On Error Resume Next
        Reg_Obj.RegDelete (Database_Path_Store)
        MsgBox "Database location cleared...", vbInformation
        End
End If
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lblBackupImport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblBackupImport, Image19
End Sub

Private Sub lblBackupImport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblBackupImport, Image19, 2
End Sub

Private Sub lblCountEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub lblCountName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub lblCourseInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblCourseInfo, Image18
End Sub

Private Sub lblCourseInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 5 Then: lblCourseInfo.ForeColor = &H6C76F2
End Sub

Private Sub lblCourseInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblCourseInfo, Image18, 2
Option_status = 5
lblCourseInfo.ForeColor = &H8000&

'lblViewExport.ForeColor = vbBlack
'lblGetEmails.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

Label10.Caption = "Enter Course Information."

lblCountEmail = "0"
lblCountName = "0"

Open_Course_Info
framCourseInfo.Visible = True
framDataViewExport.Visible = False
framEmails.Visible = False
framContacts.Visible = False
framStudentInfo.Visible = False
framOptions.Visible = False

cmdStdAdd.Default = False
cmdStdCancel.Cancel = False
cmdAdd.Default = True
cmdCancel.Cancel = True

End Sub
Private Sub lblCreateUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblCreateUser, Image14
End Sub

Private Sub lblCreateUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblCreateUser, Image14, 2
frmCreateUser.Show 1
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblExit, Image10
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 7 Then: lblExit.ForeColor = &H6C76F2
End Sub

Private Sub lblGeneral_Click()
framGeneral.Visible = True
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblExit, Image10, 2

intlogoff = 0
'Option_status = 7
'lblExit.ForeColor = &H8000&

'lblGetEmails.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack

lblCountEmail = "0"
lblCountName = "0"
Unload Me
End Sub

Private Sub lblGetContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblGetContacts, Image13
End Sub

Private Sub lblGetContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 3 Then: lblGetContacts.ForeColor = &H6C76F2
End Sub

Private Sub lblGetContacts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblGetContacts, Image13, 2

Option_status = 3
lblGetContacts.ForeColor = &H8000&

'lblGetEmails.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

Label10.Caption = "Get Contact Information of Students."

Load_Course_Info
framContacts.Visible = True
framDataViewExport.Visible = False
framEmails.Visible = False
framCourseInfo.Visible = False
framStudentInfo.Visible = False
framOptions.Visible = False

lblCountEmail = "0"
lblCountName = "0"

End Sub

Private Sub lblGetEmails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblGetEmails, Image17
End Sub

Private Sub lblGetEmails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 2 Then: lblGetEmails.ForeColor = &H6C76F2
End Sub

Private Sub lblGetEmails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblGetEmails, Image17, 2

Option_status = 2
lblGetEmails.ForeColor = &H8000&

'lblViewExport.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

lblCountEmail = "0"
lblCountName = "0"

Label10.Caption = "Get Email Information of Students."
Load_Course_Info
framEmails.Visible = True
framDataViewExport.Visible = False
framContacts.Visible = False
framCourseInfo.Visible = False
framStudentInfo.Visible = False
framOptions.Visible = False
End Sub

Private Sub lblGoToLocation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGoToLocation.Left = lblGoToLocation.Left + 20
lblGoToLocation.Top = lblGoToLocation.Top + 20

End Sub

Private Sub lblGoToLocation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGoToLocation.Left = lblGoToLocation.Left - 20
lblGoToLocation.Top = lblGoToLocation.Top - 20
Open_File_Or_Location txtDefaultFileSavingLocation.Text
End Sub

Private Sub lblLogOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblLogOff, Image5
End Sub

Private Sub lblLogOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 8 Then: lblLogOff.ForeColor = &H6C76F2
End Sub

Private Sub lblLogOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblLogOff, Image5, 2

If MsgBox("Are you sure you want to logoff active user?", vbQuestion + vbYesNo) = vbYes Then
    intPaintFlag = 0
    intlogoff = 1
    Unload frmMain
'    Unload frmStyle
    frmLogin.Show
End If

End Sub

Private Sub lblManageUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblManageUser, Image15
End Sub

Private Sub lblManageUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblManageUser, Image15, 2
frmManageUser.Show 1
End Sub

Private Sub lblOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblOptions, Image9

End Sub

Private Sub lblOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 6 Then: lblOptions.ForeColor = &H6C76F2
End Sub

Private Sub lblOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblOptions, Image9, 2

Option_status = 6

If txtImportDblocation = "" Then
    cmdImportData.Enabled = False
ElseIf txtImportDblocation <> "" Then
    cmdImportData.Enabled = True
End If

lblOptions.ForeColor = &H8000&

'lblViewExport.ForeColor = vbBlack
'lblGetEmails.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

lblCountEmail = "0"
lblCountName = "0"

On Error Resume Next
Label10.Caption = "Global Option Settings."

strMail_Client = Read_Registry(Email_Client_Loc)
strFile_Save = Read_Registry(Default_File_Loc)

txtEmailClientLocation = strMail_Client
txtDefaultFileSavingLocation = File_Root()
cmbLoadPrompt = cmbLoadPrompt.List(Read_Registry(Default_Prompt))
chkOpenFile.Value = Read_Registry(File_Open_Set)
cmbEmailSeperator = Read_Registry(Email_Seperator)

framOptions.Visible = True
framDataViewExport.Visible = False
framContacts.Visible = False
framEmails.Visible = False
framCourseInfo.Visible = False
framStudentInfo.Visible = False

End Sub

Private Sub lblProgress_Indicator_Contact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus

End Sub

Private Sub lblStudentInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblStudentInfo, Image16

End Sub

Private Sub lblStudentInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 4 Then: lblStudentInfo.ForeColor = &H6C76F2
End Sub

Private Sub lblStudentInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblStudentInfo, Image16, 2

Option_status = 4
lblStudentInfo.ForeColor = &H8000&

'lblGetEmails.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

Label10.Caption = "Enter Student Information."
Load_Course_Info
Open_Student_Info
framStudentInfo.Visible = True
framDataViewExport.Visible = False
framEmails.Visible = False
framContacts.Visible = False
framCourseInfo.Visible = False
framOptions.Visible = False

lblCountEmail = "0"
lblCountName = "0"

cmdAdd.Default = False
cmdCancel.Cancel = False
cmdStdAdd.Default = True
cmdStdCancel.Cancel = True

End Sub
Private Sub lblViewExport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblViewExport, Image12
End Sub

Private Sub lblViewExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_On_Focus
If Option_status <> 1 Then: lblViewExport.ForeColor = &H6C76F2
End Sub

Private Sub lblViewExport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Get_Control_Move lblViewExport, Image12, 2

lvwData.Sorted = False
Option_status = 1
cmdExport.Default = True
cmdCancelSearch.Cancel = True

lblViewExport.ForeColor = &H8000&

'lblGetEmails.ForeColor = vbBlack
'lblGetContacts.ForeColor = vbBlack
'lblStudentInfo.ForeColor = vbBlack
'lblCourseInfo.ForeColor = vbBlack
'lblOptions.ForeColor = vbBlack
'lblExit.ForeColor = vbBlack

lblCountEmail = "0"
lblCountName = "0"
Clear_Check_Boxes
Label10.Caption = "Data View/Export Information (Advanced)."
Load_Course_Info
framDataViewExport.Visible = True
framEmails.Visible = False
framContacts.Visible = False
framCourseInfo.Visible = False
framStudentInfo.Visible = False
framOptions.Visible = False
Load_List_View
Clear_Check_Boxes
End Sub

Private Sub lstContact_Click()
On Error Resume Next
If lstContact.Selected(lstContact.ListIndex) = True Then
    lstName.Selected(lstContact.ListIndex) = True
    'lstContact.List (lstName.ListIndex)
    'lstName.ListIndex = lstContact.ListIndex
Else
   lstName.Selected(lstContact.ListIndex) = False
End If

End Sub

Private Sub lstContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub lstEmail_Click()
Check_for_selected_email
lblCountEmail = selectedemail
End Sub

Private Sub lstEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub
'Private Sub lvwData_Click()
'Check_List_View_Selected
'End Sub

Private Sub lvwData_DblClick()
For a = 1 To lvwData.ListItems.count
    If lvwData.ListItems.Item(a).Selected = True Then
        txtSearch = lvwData.ListItems.Item(a).Text
        lblStudentInfo_MouseDown 0, 0, 0, 0
        lblStudentInfo_MouseUp 0, 0, 0, 0
        cmdSearch_Click
        Exit For
    End If
Next
' MsgBox lvwData.ListItems.Item
    'txtSearch = lvwData.ListItems(1).Text
    'lblStudentInfo_Click
    'cmdSearch_Click
'End If
End Sub

Private Sub lvwData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub lvwEmail_Click()
Check_List_View_Selected lvwEmail, 2
End Sub

Private Sub lvwEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub lvwName_Contact_Click()
Check_List_View_Selected lvwName_Contact, 1
End Sub

Private Sub lvwName_Contact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus

End Sub

Private Sub mnuEmailClipboardNormal_Click()
If Check_List_View_Selected(lvwEmail, 2) = False Then: MsgBox "Please select from the list.", vbExclamation: Exit Sub

On Error Resume Next
Clipboard.Clear
Clipboard.SetText File_Clipboard_Handle_Return_String(lvwEmail)
MsgBox "Contents successfully copied to Clipboard." & vbCrLf & "Now, you can paste it in the To: field of your Mail Application.", vbInformation
End Sub
Private Sub mnuEmailFileNormal_Click()
If Check_List_View_Selected(lvwEmail, 2) = False Then: MsgBox "Please select from the list.", vbExclamation: Exit Sub

File_Num = FreeFile

On Error Resume Next
Clipboard.Clear

If Dir(File_Root, vbDirectory) <> "" Then
    filesavepath = File_Root & "N-" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM")
    Open filesavepath & ".txt" For Output As #File_Num
         Print #File_Num, File_Clipboard_Handle_Return_String(lvwEmail)
    Close #File_Num
    
    MsgBox "File successfully saved as " & filesavepath & ".txt.", vbInformation
    
    'If strFile_Open = "" And strFile_Open <> "0" And strFile_Open <> "1" Then
    '    strFile_Open = "1"
    '    chkOpenFile.Value = 1
    '    End If
    If strFile_Open = "1" Then: Open_File_Or_Location (filesavepath & ".txt")
Else
    MsgBox "Default file saving location not found...", vbCritical
    frmBrowseFolder.Show 1
    Exit Sub
End If
End Sub

Private Sub mnuFormattedtoEmailClipboard_Click()
If Check_List_View_Selected(lvwEmail) = False Then: MsgBox "Please select from the list.", vbExclamation: Exit Sub

If strEmail_Seperator = "" Then: strEmail_Seperator = ","

Dim counter As Integer

On Error Resume Next
Clipboard.Clear
Clipboard.SetText File_Clipboard_Handle_Return_String(lvwEmail, 2)

MsgBox "Contents successfully copied to Clipboard." & vbCrLf & "Now, you can paste it in the To: field of your Mail Application.", vbInformation
End Sub

Private Sub mnuFormattedtoEmailFile_Click()
If Check_List_View_Selected(lvwEmail, 1) = False Then: MsgBox "Please select from the list.", vbExclamation: Exit Sub

File_Num = FreeFile

If strEmail_Seperator = "" Then: strEmail_Seperator = ","

On Error Resume Next
Clipboard.Clear

If Dir(File_Root, vbDirectory) <> "" Then
    filesavepath = File_Root & "F-" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM")
    Open filesavepath & ".txt" For Output As #File_Num
        Print #File_Num, File_Clipboard_Handle_Return_String(lvwEmail, 2)
    Close #File_Num
    MsgBox "File successfully saved as " & filesavepath & ".txt.", vbInformation
    'If strFile_Open = "" And strFile_Open <> "0" And strFile_Open <> "1" Then
    '     strFile_Open = "1"
    '     chkOpenFile.Value = 1
    If strFile_Open = 1 Then: Open_File_Or_Location (filesavepath & ".txt")
            
Else
    MsgBox "Default file saving location not found...", vbCritical
    frmBrowseFolder.Show 1
    Exit Sub
End If
End Sub

Private Sub mnuSaveasmsexcelfile_Click()
If Check_List_View_Selected(lvwName_Contact) = False Then: MsgBox "Please select item/s from the list.", vbExclamation: Exit Sub

On Error Resume Next
Dim intProgressLimit As Long
Dim intProChange As Integer
Dim i As Integer

Dim createExcel As New Excel.Application
Dim Wbook As Excel.Workbook
Dim Wsheet As Excel.Worksheet
Set Wbook = createExcel.Workbooks.Add
Set Wsheet = Wbook.Worksheets(1)
Wbook.Worksheets(1).Name = "Contact Information"

Wsheet.Cells(1, 1).Value = cmbSelectCourseContact
Wsheet.Cells(2, 1).Value = "Name"
Wsheet.Cells(2, 2).Value = "Contact Number"

Wsheet.Cells(1, 1).Font.Bold = True
'Wsheet.Cells(1, 1).Font.Underline = xlUnderlineStyleSingle
Wsheet.Cells(1, 1).Font.Size = 12
Wsheet.Cells(2, 1).Font.Bold = True
Wsheet.Cells(2, 2).Font.Bold = True
Wsheet.Cells(2, 1).Font.Underline = xlUnderlineStyleSingle
Wsheet.Cells(2, 2).Font.Underline = xlUnderlineStyleSingle

Dim a As String
Dim count As Integer

lngProgressLimit = lvwName_Contact.ListItems.count

lblProgress_Des_Contact.Caption = "Saving data, please wait..."
lblProgress_Indicator_Contact.Caption = "0%"

lblProgress_Des_Contact.Visible = True: lblProgress_Indicator_Contact.Visible = True

Control_Enable_With_Progress False

count = 0
For i = 1 To lvwName_Contact.ListItems.count
    If lvwName_Contact.ListItems(i).Checked = True Then
        count = count + 1
        Wsheet.Cells(count + 3, 1).Value = lvwName_Contact.ListItems(i).Text
        Wsheet.Cells(count + 3, 2).Value = lvwName_Contact.ListItems(i).ListSubItems(1).Text
    End If
    
intProgressChange = (i / lngProgressLimit) * 100
lblProgress1_Contact.Width = (intProgressChange / 100) * 4000
lblProgress2_Contact.Width = (intProgressChange / 100) * 4000
lblProgress_Indicator_Contact.Caption = intProgressChange & "%"
DoEvents
Next i

Wsheet.Columns.AutoFit
Wsheet.Cells.NumberFormat = "@"
    If Dir(File_Root, vbDirectory) <> "" Then
        filesavepath = File_Root & "Contacts-" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM")
        Wbook.SaveAs filesavepath & ".xls"
        Wbook.Close True
        lblProgress_Des_Contact.Caption = "Data saving completed..."
        Control_Enable_With_Progress True
        MsgBox "File successfully saved as " & filesavepath & ".xls.", vbInformation
        lblProgress_Des_Contact.Visible = False: lblProgress_Indicator_Contact.Visible = False
        lblProgress1_Contact.Width = 0: lblProgress2_Contact.Width = 0
        
        'If strFile_Open = "" And strFile_Open <> "0" And strFile_Open <> "1" Then
            '    strFile_Open = "1"
            '    chkOpenFile.Value = 1
            If strFile_Open = "1" Then: Open_File_Or_Location (filesavepath & ".xls")
            
    Else
        MsgBox "Default file saving location not found...", vbCritical
        frmBrowseFolder.Show 1
        Exit Sub
    End If
    
Set Wbook = Nothing
Set Wsheet = Nothing
End Sub

Private Sub mnuSaveastextfile_Click()
If Check_List_View_Selected(lvwName_Contact) = False Then: MsgBox "Please select items from the list.", vbExclamation: Exit Sub

File_Num = FreeFile

On Error Resume Next
If Dir(File_Root, vbDirectory) <> "" Then
    filesavepath = File_Root & "Contacts-" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM")
    Open filesavepath & ".txt" For Output As #File_Num
        Print #File_Num, File_Clipboard_Handle_Return_String(lvwName_Contact, 1)
        Close #File_Num
        MsgBox "File successfully saved to " & filesavepath & ".", vbInformation
        If strFile_Open = 1 Then: Open_File_Or_Location (filesavepath & ".txt")
Else
        MsgBox "Default file saving location not found...", vbCritical
        frmBrowseFolder.Show 1
        Exit Sub
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Out_Focus
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
Mouse_Out_Focus
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Picture8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then: Getmove Me
End Sub

Private Sub SelectEMclient_Click()
frmBF.Show 1
End Sub

Private Sub txtCourse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtCourseDes.SetFocus
End Sub
Private Sub txtCourseDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: cmdSave.SetFocus
End Sub

Private Sub txtDateofBirth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtStdAddress.SetFocus
End Sub

Private Sub txtDefaultFileSavingLocation_Change()
If txtDefaultFileSavingLocation <> "" Then
    lblGoToLocation.Enabled = True
Else
    lblGoToLocation.Enabled = False
End If
End Sub

Private Sub txtExamIndex_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmbCourse.SetFocus
    SendKeys "{F4}"
End If
End Sub

Private Sub txtImportDblocation_Change()
If txtImportDblocation = "" Then
    cmdImportData.Enabled = False
ElseIf txtImportDblocation <> "" Then
    cmdImportData.Enabled = True
End If
End Sub

Private Sub txtNICPassport_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtStdName.SetFocus
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then: cmdStdSave.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
If intlogoff = 0 Then
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbYes Then
        'Cancel = 0
        dbcon.Close
        Set dbcon = Nothing
        Set Me.Picture = Nothing
        Unload frmSplash
        End
    Else
        Cancel = 1
        'lblViewExport_MouseUp 0, 0, 0, 0
    End If
End If
End Sub

Private Sub lblEx_Pro_Description_Click()
Open_Course_Info
framCourseInfo.Visible = True
framStudentInfo.Visible = False
End Sub
Private Sub Label7_Click()
Unload Me
End Sub
Private Sub cmdAdd_Click()
'Enable_Controls
If Privilege_Proceed = False Then: Exit Sub
blnadd = True
Clear_Fields_Course
txtCourse.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode_Course False
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
rstcourseinfo.CancelUpdate
Clear_Fields_Course
Inforfield_Course
Disable_Controls_Course
blnadd = False
blnedit = False
End Sub

Private Sub cmdDelete_Click()
'blndelete_Click = True
If Privilege_Proceed = False Then: Exit Sub
If MsgBox("Delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
    On Error Resume Next
    rstcourseinfo.Delete
    rstcourseinfo.MoveNext
        If rstcourseinfo.EOF Then
            rstcourseinfo.MoveLast
        End If
    Inforfield_Course
End If
End Sub

Private Sub cmdEdit_Click()
If Privilege_Proceed = False Then: Exit Sub
firstcourse = txtCourse
blnedit = True
Enable_Controls_Course
txtCourse.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode_Course False
End Sub

Private Sub cmdFirst_Click()
   If rstcourseinfo.BOF = False Then
        rstcourseinfo.MoveFirst
        Inforfield_Course
        MsgBox "You are on the First Record.", vbInformation
    End If
End Sub

Private Sub cmdLast_Click()
  If rstcourseinfo.EOF = False Then
        rstcourseinfo.MoveLast
        Inforfield_Course
        MsgBox "You are on the Last Record.", vbInformation
    End If
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
   If rstcourseinfo.EOF = False Then
        rstcourseinfo.MoveNext
        Inforfield_Course
    End If
    If rstcourseinfo.EOF Then
        rstcourseinfo.MoveLast
        Inforfield_Course
        MsgBox "You are on the Last Record.", vbInformation
    End If
End Sub

Private Sub cmdPrevious_Click()
   If rstcourseinfo.BOF = False Then
        rstcourseinfo.MovePrevious
        Inforfield_Course
   ElseIf rstcourseinfo.BOF = True Then
        rstcourseinfo.MoveFirst
        Inforfield_Course
        MsgBox "You are on the first Record.", vbInformation
    End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim rstRecordexistcourse As ADODB.Recordset
Dim Find_recordexist_course As String
If blnadd = True Then
    If txtCourse = "" Then: MsgBox "Field cannot be blank...!", vbExclamation: txtCourse.SetFocus: Exit Sub
    Find_recordexist_course = txtCourse
    If Check_for_Record_Existence("COURSEINFO", "COURSE", Find_recordexist_course) = True Then: txtCourse.SelStart = 0: txtCourse.SelLength = Len(txtCourse): txtCourse.SetFocus: Exit Sub
    rstcourseinfo.AddNew
    Save_Data_Course
ElseIf blnedit = True Then
    If txtCourse = "" Then: MsgBox "Field cannot be blank...!", vbExclamation: txtCourse.SetFocus: Exit Sub
    secondcourse = txtCourse
        If firstcourse <> secondcourse Then
            Find_recordexist_course = txtCourse
            If Check_for_Record_Existence("COURSEINFO", "COURSE", Find_recordexist_course) = True Then: txtCourse.SelStart = 0: txtCourse.SelLength = Len(txtCourse): txtCourse.SetFocus: Exit Sub
        End If
    Save_Data_Course
End If
End Sub

Public Sub Inforfield_Course()
On Error Resume Next
If rstcourseinfo.RecordCount > 0 Then
    txtCourse = rstcourseinfo("COURSE")
    txtCourseDes = rstcourseinfo("COURSE_DES")
    Button_Add_Edit_Save_Cancle_RecordExist_Mode_Course True
Else
    Clear_Fields_Course
    Disable_Controls_Course
    Button_Record_Not_Exist_Mode_Course
    Exit Sub
End If
End Sub

Sub Save_Data_Course()

rstcourseinfo("COURSE") = txtCourse
rstcourseinfo("COURSE_DES") = Format_Data_Field(Trim(txtCourseDes))
rstcourseinfo.Update

Button_Add_Edit_Save_Cancle_RecordExist_Mode_Course True
Disable_Controls_Course
blnadd = False
blnedit = False
End Sub

Public Sub Button_Add_Edit_Save_Cancle_RecordExist_Mode_Course(bval As Boolean)
cmdCancel.Enabled = Not bval: cmdSave.Enabled = Not bval
cmdAdd.Enabled = bval: cmdEdit.Enabled = bval
cmdDelete.Enabled = bval: cmdPrevious.Enabled = bval
cmdNext.Enabled = bval: cmdFirst.Enabled = bval
cmdLast.Enabled = bval
End Sub

Public Sub Button_Record_Not_Exist_Mode_Course()
cmdAdd.Enabled = True: cmdEdit.Enabled = False
cmdDelete.Enabled = False: cmdCancel.Enabled = False
cmdSave.Enabled = False: cmdPrevious.Enabled = False
cmdNext.Enabled = False: cmdFirst.Enabled = False
cmdLast.Enabled = False
End Sub
Public Sub Clear_Fields_Course()
Enable_Controls_Course
txtCourse = ""
txtCourseDes = ""
End Sub
Public Sub Disable_Controls_Course()
txtCourse.Enabled = False
txtCourseDes.Enabled = False
End Sub

Public Sub Enable_Controls_Course()
txtCourse.Enabled = True
txtCourseDes.Enabled = True
End Sub

Public Sub Open_Student_Info()
Set rststdinfo = New ADODB.Recordset
    rststdinfo.CursorLocation = adUseClient
    rststdinfo.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
    
Inforfield
Disable_Controls
cmdCancelSearch.Enabled = False
cmbSearchOption.Clear
cmbSearchOption.AddItem "Student ID"
cmbSearchOption.AddItem "Exam Index"
cmbSearchOption.AddItem "NIC/Passport"
cmbSearchOption.AddItem "Name"
cmbSearchOption = cmbSearchOption.List(0)

'cmbSearchOption.AddItem "Student ID"

'cmbSearchCriteria.Clear
'cmbSearchCriteria.AddItem "Student ID"
'cmbSearchCriteria.AddItem "Exam Index"
'cmbSearchCriteria.AddItem "NIC/Passport"
'cmbSearchCriteria.AddItem "Name"
'cmbSearchCriteria.AddItem "By Name"
'cmbSearchOption.AddItem "Student ID"
'cmbSearchCriteria = cmbSearchCriteria.List(0)
End Sub
Public Sub Open_Course_Info()
Set rstcourseinfo = New ADODB.Recordset
    rstcourseinfo.CursorLocation = adUseClient
    rstcourseinfo.Open "SELECT * FROM COURSEINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
Inforfield_Course
Disable_Controls_Course
End Sub
Public Sub Load_Course_Info()
On Error Resume Next
Dim rstloadcourseinfo As ADODB.Recordset
Set rstloadcourseinfo = New ADODB.Recordset
    rstloadcourseinfo.CursorLocation = adUseClient
    rstloadcourseinfo.Open "SELECT * FROM COURSEINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
    cmbCourse.Clear
    cmbSelectCourseContact.Clear
    cmbSelectCourseEmail.Clear
    cmbSelectCourseExport.Clear
    cmbSelectCourseContact.AddItem "---All Courses---"
    cmbSelectCourseEmail.AddItem "---All Courses---"
    cmbSelectCourseExport.AddItem "---All Courses---"
        If rstloadcourseinfo.RecordCount > 0 Then
            Do While Not rstloadcourseinfo.EOF
                cmbCourse.AddItem rstloadcourseinfo("COURSE")
                cmbSelectCourseContact.AddItem rstloadcourseinfo("COURSE")
                cmbSelectCourseEmail.AddItem rstloadcourseinfo("COURSE")
                cmbSelectCourseExport.AddItem rstloadcourseinfo("COURSE")
                rstloadcourseinfo.MoveNext
            Loop
        End If
'cmbCourse.Text = cmbCourse.List(0)
cmbSelectCourseContact = cmbSelectCourseContact.List(0)
cmbSelectCourseEmail = cmbSelectCourseEmail.List(0)
cmbSelectCourseExport = cmbSelectCourseExport.List(0)

cmbSex.Clear
cmbSex.AddItem "Male"
cmbSex.AddItem "Female"

rstloadcourseinfo.Close
Set rstloadcourseinfo = Nothing
End Sub
Public Sub Check_for_selected_email()
selectedemail = 0
For s = 0 To lstEmail.ListCount - 1
    If lstEmail.Selected(s) = True Then
        selectedemail = selectedemail + 1
    End If
Next s
End Sub
Public Sub Load_Options_Info()
cmbLoadPrompt.AddItem "Data View/Export"
cmbLoadPrompt.AddItem "Get Emails"
cmbLoadPrompt.AddItem "Get Contact"
cmbLoadPrompt.AddItem "Student Info"
cmbLoadPrompt.AddItem "Coruse Info"
cmbLoadPrompt.AddItem "Options"
cmbLoadPrompt = cmbLoadPrompt.List(0)
cmbEmailSeperator.AddItem ", - comma"
cmbEmailSeperator.AddItem "; - semicolon"
cmbEmailSeperator = cmbEmailSeperator.List(0)
cmbClearDatabase.AddItem "Course Information"
cmbClearDatabase.AddItem "Student Information"
cmbClearDatabase = cmbClearDatabase.List(0)

If intaccount_type = 0 Then
    cmdSelectDatabase.Enabled = False
    cmdImportData.Enabled = False
    cmdClear.Enabled = False
    cmbClearDatabase.Enabled = False
    'cmdFormatDatabase.Enabled = False
    'cmdFormatDatabaseDV.Enabled = False
Else
    cmdSelectDatabase.Enabled = True
    cmdImportData.Enabled = True
    cmdClear.Enabled = True
    cmbClearDatabase.Enabled = True
    'cmdFormatDatabase.Enabled = True
    'cmdFormatDatabaseDV.Enabled = True
End If

If user_write_privilege = 0 Then
    cmdClear.Enabled = False: cmdFormatDatabase.Enabled = False
    cmdSelectDatabase.Enabled = False: cmdImportData.Enabled = False
    cmbClearDatabase.Enabled = False: cmdFormatDatabaseDV.Enabled = False
End If

On Error Resume Next
strMail_Client = Read_Registry(Email_Client_Loc)
strFile_Save = Read_Registry(Default_File_Loc)
txtEmailClientLocation = strMail_Client
txtDefaultFileSavingLocation = strFile_Save
cmbLoadPrompt = cmbLoadPrompt.List(Read_Registry(Default_Prompt))
cmbEmailSeperator = Read_Registry(Email_Seperator)
chkOpenFile.Value = Read_Registry(File_Open_Set)
chkEnablePassword.Value = Read_Registry(Pword)
End Sub
Public Sub Mouse_On_Focus()
Select Case Option_status
    Case 1: lblViewExport.ForeColor = &H8000&: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 2: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = &H8000&
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 3: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = &H8000&: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 4: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = &H8000&
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 5: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = &H8000&: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 6: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = &H8000&
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 7: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = &H8000&: lblLogOff.ForeColor = vbBlack
    Case 8: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = &H8000&: lblLogOff.ForeColor = &H8000&
End Select
End Sub

Public Sub Mouse_Out_Focus()
Select Case Option_status
    Case 1: lblViewExport.ForeColor = &H8000&: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 2: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = &H8000&
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 3: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = &H8000&: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 4: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = &H8000&
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 5: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = &H8000&: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 6: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = &H8000&
            lblExit.ForeColor = vbBlack: lblLogOff.ForeColor = vbBlack
    Case 7: lblViewExport.ForeColor = vbBlack: lblGetEmails.ForeColor = vbBlack
            lblGetContacts.ForeColor = vbBlack: lblStudentInfo.ForeColor = vbBlack
            lblCourseInfo.ForeColor = vbBlack: lblOptions.ForeColor = vbBlack
            lblExit.ForeColor = &H8000&: lblLogOff.ForeColor = vbBlack
    End Select
End Sub

Private Sub txtSearchDataView_Change()
'If txtSearchDataView.Text = "" Then
'    MsgBox "Search field is empty...!", vbExclamation
'    txtSearchDataView.SetFocus
'    Exit Sub
'End If

'Set rststdinfosearch = New ADODB.Recordset
    'rststdinfosearch.CursorLocation = adUseClient
    'rststdinfosearch.Open "SELECT * FROM CusInfo ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
On Error Resume Next
Find_val = txtSearchDataView
Set rststdinfodv = New ADODB.Recordset

If cmbSelectCourseExport = "---All Courses---" Then
    Select Case cmbSearchCriteria.ListIndex
        Case 0: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_ID] Like '" & Find_val & "%'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 1: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_EXAM_INDEX] Like '" & Find_val & "%'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 2: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_NIC_PASSPORT] Like '" & Find_val & "%'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 3: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_NAME] Like '" & Find_val & "%'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 4: intAdvanced_search = 1: rststdinfodv.Open "SELECT* FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
    End Select
Else
    Select Case cmbSearchCriteria.ListIndex
        Case 0: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_ID] Like '" & Find_val & "%' AND STD_COURSE = '" & cmbSelectCourseExport & "'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 1: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_EXAM_INDEX] Like '" & Find_val & "%' AND STD_COURSE = '" & cmbSelectCourseExport & "'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 2: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_NIC_PASSPORT] Like '" & Find_val & "%' AND STD_COURSE = '" & cmbSelectCourseExport & "'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 3: intAdvanced_search = 0: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_NAME] Like '" & Find_val & "%' AND STD_COURSE = '" & cmbSelectCourseExport & "'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
        Case 4: intAdvanced_search = 1: rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_COURSE] = '" & cmbSelectCourseExport & "'", dbcon, adOpenStatic, adLockOptimistic: Load_to_List
    End Select
    rststdinfodv.Close
    Set rststdinfodv = Nothing
End If
Clear_Check_Boxes
chkSelectAllItems.Value = 0
lblSelectedCountDV.Caption = "0"
End Sub

Private Sub txtStdAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtStdTel.SetFocus
End Sub

Private Sub txtStdEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtRemarks.SetFocus
End Sub

Private Sub txtStdID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtExamIndex.SetFocus
End Sub

Private Sub txtStdName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: cmbSex.SetFocus
End Sub

Private Sub txtStdTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtStdEmail.SetFocus
End Sub
Private Sub chkSelectAllItems_Click()
If chkSelectAllItems.Value = 1 Then
    For i = 1 To lvwData.ListItems.count
    lvwData.ListItems(i).Checked = True
    Check_List_View_Selected lvwData
    Next
Else
    For i = 1 To lvwData.ListItems.count
    lvwData.ListItems(i).Checked = False
    lblSelectedCountDV.Caption = "0"
    Next
End If
End Sub

Private Sub cmbSelectCourseExport_Click()
Dim selectedcoursefordv As String
Dim i As ListItem
lvwData.SmallIcons = imgList
lvwData.ListItems.Clear
Clear_Check_Boxes
lvwData.Sorted = False

On Error Resume Next

txtSearchDataView = ""
chkSelectAllItems.Value = 0
lblSelectedCountDV.Caption = "0"
If cmbSelectCourseExport = "---All Courses---" Then

Set rststdinfodv = New ADODB.Recordset
    rststdinfodv.CursorLocation = adUseClient
    rststdinfodv.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
        If rststdinfodv.RecordCount > 0 Then
            Do While Not rststdinfodv.EOF
                Set i = lvwData.ListItems.Add(, , rststdinfodv("STD_ID"))
                    i.SubItems(1) = rststdinfodv("STD_EXAM_INDEX")
                    i.SubItems(2) = rststdinfodv("STD_COURSE")
                    i.SubItems(3) = rststdinfodv("STD_NIC_PASSPORT")
                    i.SubItems(4) = rststdinfodv("STD_NAME")
                    i.SubItems(5) = rststdinfodv("SEX")
                    i.SubItems(6) = rststdinfodv("DATE_OF_BIRTH")
                    i.SubItems(7) = rststdinfodv("STD_ADDRESS")
                    i.SubItems(8) = rststdinfodv("STD_CONTACT")
                    i.SubItems(9) = rststdinfodv("STD_EMAIL")
                    i.SubItems(10) = rststdinfodv("REMARKS")
                    i.SmallIcon = 1
                rststdinfodv.MoveNext
            Loop
        End If
 
Else
    selectedcoursefordv = cmbSelectCourseExport
    Set rststdinfodv = New ADODB.Recordset
        rststdinfodv.Open "SELECT* FROM STUDENTINFO WHERE[STD_COURSE] = '" & selectedcoursefordv & "' ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
            If rststdinfodv.RecordCount > 0 Then
                Do While Not rststdinfodv.EOF
                    Set i = lvwData.ListItems.Add(, , rststdinfodv("STD_ID"))
                        i.SubItems(1) = rststdinfodv("STD_EXAM_INDEX")
                        i.SubItems(2) = rststdinfodv("STD_COURSE")
                        i.SubItems(3) = rststdinfodv("STD_NIC_PASSPORT")
                        i.SubItems(4) = rststdinfodv("STD_NAME")
                        i.SubItems(5) = rststdinfodv("SEX")
                        i.SubItems(6) = rststdinfodv("DATE_OF_BIRTH")
                        i.SubItems(7) = rststdinfodv("STD_ADDRESS")
                        i.SubItems(8) = rststdinfodv("STD_CONTACT")
                        i.SubItems(9) = rststdinfodv("STD_EMAIL")
                        i.SubItems(10) = rststdinfodv("REMARKS")
                        i.SmallIcon = 1
                        rststdinfodv.MoveNext
                Loop
            End If
End If
rststdinfodv.Close
Set rststdinfodv = Nothing
End Sub
Private Sub cmdExport_Click()
'Reset_Progress
If Check_List_View_Selected(lvwData) = False Then: MsgBox "Please select items from the list.", vbExclamation: Exit Sub
If Checkl_Fields_Selected = False Then MsgBox "Please select fields to be exported.", vbExclamation: Exit Sub
Data_Export
End Sub

Private Sub cmdIncluedAll_Click()
For i = 0 To 10
    chkExportField.Item(i).Value = 1
Next
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvwData.SortKey = ColumnHeader.Index - 1
lvwData.Sorted = True
End Sub
Public Sub Data_Export()
Dim intProgressLimit As Long
Dim intProChange As Integer

Reset_Progress 2
Control_Enable_With_Progress False

On Error Resume Next
Dim createExcel As New Excel.Application
Dim Wbook As Excel.Workbook
Dim Wsheet As Excel.Worksheet
Set Wbook = createExcel.Workbooks.Add
Set Wsheet = Wbook.Worksheets(1)
Wbook.Worksheets(1).Name = "Student Info"

Wsheet.Cells(1, 1).Value = cmbSelectCourseExport
Wsheet.Cells(1, 1).Font.Bold = True
Wsheet.Cells(1, 1).Font.Size = 11

Dim fcount, lcount As Integer
fcount = 0
For i = 0 To 10
    If chkExportField(i).Value = 1 Then
        fcount = fcount + 1
        Wsheet.Cells(2, fcount).Value = chkExportField(i).Caption
        Wsheet.Cells(2, fcount).Font.Bold = True
        Wsheet.Cells(2, fcount).Font.Size = 10
        Wsheet.Cells(2, fcount).Font.Underline = xlUnderlineStyleSingle
    End If
Next

lngProgressLimit = lvwData.ListItems.count
lcount = 0
For l = 1 To lvwData.ListItems.count
    If lvwData.ListItems(l).Checked = True Then
        lcount = lcount + 1
        fcount = 0
        For i = 1 To 10
            If chkExportField(i).Value = 1 Then
                fcount = fcount + 1
                Wsheet.Cells(lcount + 3, 1).Value = lvwData.ListItems(l).Text
                Wsheet.Cells(lcount + 3, fcount + 1).Value = lvwData.ListItems(l).ListSubItems(i).Text
            End If
        Next
    End If
    
intProgressChange = (l / lngProgressLimit) * 100
lblProgress1.Width = (intProgressChange / 100) * 4000
lblProgress2.Width = (intProgressChange / 100) * 4000
lblPro_Indicator.Caption = intProgressChange & "%"
DoEvents

Next

If intProChange = "100%" Then: lblEx_Pro_Description.Caption = "Data Export completed..."
Wsheet.Columns.AutoFit
Wsheet.Rows.AutoFit
Wsheet.Cells.NumberFormat = "@"
If Dir(File_Root, vbDirectory) <> "" Then
    filesavepath = File_Root & "StudentInfo" & Format(Date, "dd-mm-yyyy") & " - " & Format(Time, "hh.mm.ss AMPM") & ".xls"
    Wbook.SaveAs filesavepath
    Wbook.Close True
    
    Control_Enable_With_Progress True
    MsgBox "File successfully saved to " & filesavepath & ".", vbInformation
    Reset_Progress 1
    
    'If strFile_Open = "" And strFile_Open <> "0" And strFile_Open <> "1" Then
    '    strFile_Open = "1"
    '    chkOpenFile.Value = 1
    '    ShellExecute hwnd, "Open", filesavepath, vbNullString, vbNullString, SW_SHOW
    If strFile_Open = "1" Then: Open_File_Or_Location (filesavepath)
   
Else
    MsgBox "Default file saving location not found...", vbCritical
    frmBrowseFolder.Show 1
    Exit Sub
End If

Set Wbook = Nothing
Set Wsheet = Nothing
End Sub


Public Function Check_List_View_Selected(ByVal Required_List As ListView, Optional C_Opt As Integer = 0) As Boolean
Dim selectedlistitem As Integer

Check_List_View_Selected = False
selectedlistitem = 0

For i = 1 To Required_List.ListItems.count
    If Required_List.ListItems(i).Checked = True Then
        selectedlistitem = selectedlistitem + 1
    End If
Next
If selectedlistitem > 0 Then: Check_List_View_Selected = True

Select Case C_Opt
    Case 0: lblSelectedCountDV.Caption = selectedlistitem
    Case 1: lblCountName.Caption = selectedlistitem
    Case 2: lblCountEmail.Caption = selectedlistitem
End Select
End Function

Public Function Checkl_Fields_Selected() As Boolean
Dim selectedfields As Integer

selectedfields = 0
Checkl_Fields_Selected = False

For i = 0 To 10
    If chkExportField(i).Value = 1 Then
        selectedfields = selectedfields + 1
    End If
Next
If selectedfields > 1 Then: Checkl_Fields_Selected = True
End Function

Public Sub Load_to_List()
On Error Resume Next
lvwData.ListItems.Clear
If intAdvanced_search = 0 Then
    If rststdinfodv.RecordCount > 0 Then
        Do While Not rststdinfodv.EOF
            Set i = lvwData.ListItems.Add(, , rststdinfodv("STD_ID"))
                i.SubItems(1) = rststdinfodv("STD_EXAM_INDEX")
                i.SubItems(2) = rststdinfodv("STD_COURSE")
                i.SubItems(3) = rststdinfodv("STD_NIC_PASSPORT")
                i.SubItems(4) = rststdinfodv("STD_NAME")
                i.SubItems(5) = rststdinfodv("SEX")
                i.SubItems(6) = rststdinfodv("DATE_OF_BIRTH")
                i.SubItems(7) = rststdinfodv("STD_ADDRESS")
                i.SubItems(8) = rststdinfodv("STD_CONTACT")
                i.SubItems(9) = rststdinfodv("STD_EMAIL")
                i.SubItems(10) = rststdinfodv("REMARKS")
                i.SmallIcon = 1
                rststdinfodv.MoveNext
                'DoEvents
        Loop
    End If
ElseIf intAdvanced_search = 1 Then
    If rststdinfodv.RecordCount > 0 Then
        Do While Not rststdinfodv.EOF
            If InStr(1, rststdinfodv("STD_NAME"), Find_val, vbTextCompare) Then
                Set i = lvwData.ListItems.Add(, , rststdinfodv("STD_ID"))
                    i.SubItems(1) = rststdinfodv("STD_EXAM_INDEX")
                    i.SubItems(2) = rststdinfodv("STD_COURSE")
                    i.SubItems(3) = rststdinfodv("STD_NIC_PASSPORT")
                    i.SubItems(4) = rststdinfodv("STD_NAME")
                    i.SubItems(5) = rststdinfodv("SEX")
                    i.SubItems(6) = rststdinfodv("DATE_OF_BIRTH")
                    i.SubItems(7) = rststdinfodv("STD_ADDRESS")
                    i.SubItems(8) = rststdinfodv("STD_CONTACT")
                    i.SubItems(9) = rststdinfodv("STD_EMAIL")
                    i.SubItems(10) = rststdinfodv("REMARKS")
                    i.SmallIcon = 1
            End If
            rststdinfodv.MoveNext
            'DoEvents
        Loop
    End If
End If

rststdinfodv.Close
Set rststdinfodv = Nothing
End Sub

Public Sub Load_List_View()
lvwData.ListItems.Clear
lvwData.SmallIcons = imgList

On Error Resume Next

Set rststdinfodv = New ADODB.Recordset
    rststdinfodv.CursorLocation = adUseClient
    rststdinfodv.Open "SELECT * FROM STUDENTINFO ORDER BY REC_ID", dbcon, adOpenStatic, adLockOptimistic
    
Dim i As ListItem

Do While Not rststdinfodv.EOF
    Set i = lvwData.ListItems.Add(, , rststdinfodv("STD_ID"))
        i.SubItems(1) = rststdinfodv("STD_EXAM_INDEX")
        i.SubItems(2) = rststdinfodv("STD_COURSE")
        i.SubItems(3) = rststdinfodv("STD_NIC_PASSPORT")
        i.SubItems(4) = rststdinfodv("STD_NAME")
        i.SubItems(5) = rststdinfodv("SEX")
        i.SubItems(6) = rststdinfodv("DATE_OF_BIRTH")
        i.SubItems(7) = rststdinfodv("STD_ADDRESS")
        i.SubItems(8) = rststdinfodv("STD_CONTACT")
        i.SubItems(9) = rststdinfodv("STD_EMAIL")
        i.SubItems(10) = rststdinfodv("REMARKS")
        i.SmallIcon = 1
    rststdinfodv.MoveNext
    'If intPaintFlag = 1 Then: DoEvents
Loop

cmdCancelSearch.Enabled = False
cmbSearchCriteria.Clear
cmbSearchCriteria.AddItem "Student ID"
cmbSearchCriteria.AddItem "Exam Index"
cmbSearchCriteria.AddItem "NIC/Passport"
cmbSearchCriteria.AddItem "Name"
cmbSearchCriteria.AddItem "Name - Advanced"
'cmbSearchOption.AddItem "Student ID"
cmbSearchCriteria = cmbSearchCriteria.List(0)
End Sub

Public Sub Clear_Check_Boxes()
For n = 1 To 10
    chkExportField(n).Value = 0
Next
End Sub

Private Sub Get_Control_Move(ByVal Label_Name As Label, ByVal Image_Name As Image, Optional opt As Integer = 1)
Select Case opt
    Case 1: Label_Name.Top = Label_Name.Top + 20: Label_Name.Left = Label_Name.Left + 20
            Image_Name.Top = Image_Name.Top + 20: Image_Name.Left = Image_Name.Left + 20
    Case 2: Label_Name.Top = Label_Name.Top - 20: Label_Name.Left = Label_Name.Left - 20
            Image_Name.Top = Image_Name.Top - 20: Image_Name.Left = Image_Name.Left - 20
    End Select
End Sub
Public Function Privilege_Proceed() As Boolean
    Privilege_Proceed = True
If user_write_privilege = 0 Then: MsgBox "You do not have permission to proceed." & vbCrLf & "Contact Administrator.", vbCritical: Privilege_Proceed = False
End Function

Public Sub Check_Privilege_Commands()
If user_write_privilege = 0 Then
    cmdClear.Enabled = False: cmdFormatDatabase.Enabled = False
    cmdSelectDatabase.Enabled = False: cmdImportData.Enabled = False
    cmbClearDatabase.Enabled = False
End If
End Sub
Public Sub Reset_Progress(opt As Integer)
Select Case opt
    Case 1: lblProgress1.Width = 0: lblProgress2.Width = 0
            lblEx_Pro_Description.Caption = "Exporting data, please wait..."
            lblPro_Indicator.Caption = "0%": 'cmdExport.Enabled = True
            lblEx_Pro_Description.Visible = False: lblPro_Indicator.Visible = False
    Case 2: lblEx_Pro_Description.Visible = True
            lblPro_Indicator.Visible = True
            'cmdExport.Enabled = False
    Case 3: lblProgressImport1.Width = 0: lblProgressImport2.Width = 0
            lblImportDescription.Caption = "Importing data, please wait..."
            lblImportIndicator.Caption = "%0": lblImportIndicator.Visible = False
            lblImportDescription.Visible = False: 'cmdClear.Enabled = True
            'cmdFormatDatabase.Enabled = True: cmdSelectDatabase.Enabled = True
            'cmdImportData.Enabled = True
    Case 4: lblImportIndicator.Visible = True
            lblImportDescription.Visible = True
            'cmdClear.Enabled = False: cmdFormatDatabase.Enabled = False
            'cmdSelectDatabase.Enabled = False: cmdImportData.Enabled = False
    Case 5: lblProgress1.Width = 0: lblProgress2.Width = 0
            lblProgressImport1.Width = 0: lblProgressImport2.Width = 0
            'lblEx_Pro_Description.Caption = "Formating data, please wait..."
            lblPro_Indicator.Caption = "0%": 'cmdExport.Enabled = True
            'cmdFormatDatabaseDV.Enabled = True: cmdFormatDatabase.Enabled = True
            lblEx_Pro_Description.Visible = False: lblPro_Indicator.Visible = False
            lblImportDescription.Visible = False: lblImportIndicator.Visible = False
            lblProgress1.BackColor = &HB3F9B8: lblProgress2.BackColor = &H1BD805
            lblProgressImport1.BackColor = &H1BD805: lblProgressImport2.BackColor = &HB3F9B8
    Case 6: lblEx_Pro_Description.Caption = "Formating data, please wait..."
            lblImportDescription.Caption = "Formating data, please wait..."
            lblProgress1.BackColor = &HC0C0FF: lblProgress2.BackColor = &H4759FE
            lblProgressImport1.BackColor = &H4759FE: lblProgressImport2.BackColor = &HC0C0FF
            lblEx_Pro_Description.Visible = True: lblPro_Indicator.Visible = True
            lblImportIndicator.Visible = True: lblImportDescription.Visible = True
            'cmdExport.Enabled = False: cmdFormatDatabaseDV.Enabled = False
            'cmdFormatDatabase.Enabled = False
            lblProgress1.Refresh: lblProgress2.Refresh
    Case 7: lblImportDescription.Caption = "Deleting data, please wait..."
            lblPro_Indicator.Visible = True: lblImportIndicator.Visible = True
            lblImportDescription.Visible = True: 'cmdExport.Enabled = False
            'cmdFormatDatabase.Enabled = False
End Select
End Sub

Public Function File_Root() As String
File_Root = strFile_Save
If File_Root = "" Or Dir(File_Root, vbDirectory) = "" Then
    File_Root = App.Path
        If Right(App.Path, 1) <> "\" Then
            File_Root = App.Path + "\"
        Else
            File_Root = App.Path
        End If
End If
End Function
Public Function Format_Data_Field(ByVal strFieldData As String) As String
'This function is used to remove unnecessary characters like New Line characters,(which may add to data by pressing Enter key)
'from a given string and format the string so that it can not contain such characters. This process searches for such characters
'in the given string from beginning to the end and removes them, if found.

Dim strChkStr As String
Dim strFormatedStr As String
Dim strCharacter As String
Dim i As Integer

strChkStr = Trim(strFieldData)
strFormatedStr = ""
If Len(strChkStr) > 0 Then
    For i = 1 To Len(strChkStr)
        If Asc(Mid(strChkStr, i, 1)) = 13 Or Asc(Mid(strChkStr, i, 1)) = 10 Then
            strCharacter = ""
        Else
            strCharacter = Mid(strChkStr, i, 1)
        End If
                strFormatedStr = strFormatedStr & strCharacter
    Next
    Format_Data_Field = Trim(strFormatedStr)
End If
End Function
Public Function File_Clipboard_Handle_Return_String(ByVal ListV As ListView, Optional opt As Integer = 0) As String
Dim i, counter As Integer
strAppend = ""
Select Case opt
Case 0
        For i = 1 To ListV.ListItems.count
            If ListV.ListItems(i).Checked = True Then
                strAppend = strAppend + ListV.ListItems(i).Text & vbCrLf
            End If
        Next
Case 1
        For i = 1 To ListV.ListItems.count
            If ListV.ListItems(i).Checked = True Then
                strAppend = strAppend + ListV.ListItems(i).Text & " ---------------> " & _
                ListV.ListItems(i).ListSubItems(1).Text & vbCrLf
            End If
        Next

Case 2
        For i = 1 To ListV.ListItems.count
            If ListV.ListItems(i).Checked = True Then
                counter = counter + 1
                    If counter = lblCountEmail.Caption Then
                        strAppend = strAppend + ListV.ListItems(i).Text
                    Else
                        strAppend = strAppend + ListV.ListItems(i).Text & strEmail_Seperator
                    End If
            End If
        Next
End Select
File_Clipboard_Handle_Return_String = strAppend
End Function
Public Sub Open_File_Or_Location(ByVal F_Loc As String)
ShellExecute hwnd, "Open", F_Loc, vbNullString, vbNullString, SW_SHOW
End Sub
