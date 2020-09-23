VERSION 5.00
Begin VB.Form frmSplash 
   Caption         =   "Transparancy demo splash form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TransparancyDemo.FormTransparancy FormTransparancy1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   873
      _ExtentY        =   873
      TransparencyDirection=   2
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
  
  If FormTransparancy1.TransparencyLevel > 0 Then
    Cancel = True
    Me!FormTransparancy1.TransparencyDirection = -2
  End If
  
End Sub
