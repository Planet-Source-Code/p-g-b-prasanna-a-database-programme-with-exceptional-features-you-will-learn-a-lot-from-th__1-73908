VERSION 5.00
Begin VB.Form frmBG_Image 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgBG_Image 
      Height          =   600
      Left            =   0
      Picture         =   "frmBG_Image.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmBG_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
