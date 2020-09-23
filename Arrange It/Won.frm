VERSION 5.00
Begin VB.Form Won 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Won.frx":0000
   ScaleHeight     =   1620
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Won"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

