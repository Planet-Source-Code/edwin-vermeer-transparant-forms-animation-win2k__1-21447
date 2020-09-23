VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00000000&
   Caption         =   "Transparency demo"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin TransparancyDemo.FormTransparancy FormTransparancy1 
      Left            =   600
      Top             =   2520
      _ExtentX        =   661
      _ExtentY        =   661
      TransparencyLevel=   100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show splash screen"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   1680
      Picture         =   "frmDemo.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1890
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  frmSplash.Show
End Sub
