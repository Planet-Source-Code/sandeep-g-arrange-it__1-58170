VERSION 5.00
Begin VB.Form HlpFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "HlpFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed By: Sandeep.G"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"HlpFrm.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrange't"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "HlpFrm.frx":0431
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "HlpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
    HlpFrm.SetFocus
End Sub
