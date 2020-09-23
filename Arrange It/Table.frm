VERSION 5.00
Begin VB.Form Table 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrange't"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3930
   Icon            =   "Table.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      Picture         =   "Table.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   24
      Left            =   6360
      Picture         =   "Table.frx":1C4C
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   23
      Left            =   5760
      Picture         =   "Table.frx":2F4E
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   22
      Left            =   5160
      Picture         =   "Table.frx":4250
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   21
      Left            =   4560
      Picture         =   "Table.frx":5552
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   20
      Left            =   3960
      Picture         =   "Table.frx":6854
      Top             =   5760
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   19
      Left            =   6360
      Picture         =   "Table.frx":7B56
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   18
      Left            =   5760
      Picture         =   "Table.frx":8E58
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   17
      Left            =   5160
      Picture         =   "Table.frx":A15A
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   16
      Left            =   4560
      Picture         =   "Table.frx":B45C
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   15
      Left            =   3960
      Picture         =   "Table.frx":C75E
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   14
      Left            =   6360
      Picture         =   "Table.frx":DA60
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   13
      Left            =   5760
      Picture         =   "Table.frx":ED62
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   12
      Left            =   5160
      Picture         =   "Table.frx":10064
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   11
      Left            =   4560
      Picture         =   "Table.frx":11366
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   10
      Left            =   3960
      Picture         =   "Table.frx":12668
      Top             =   4560
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   9
      Left            =   6360
      Picture         =   "Table.frx":1396A
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   8
      Left            =   5760
      Picture         =   "Table.frx":14C6C
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   7
      Left            =   5160
      Picture         =   "Table.frx":15F6E
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   6
      Left            =   4560
      Picture         =   "Table.frx":17270
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   5
      Left            =   3960
      Picture         =   "Table.frx":18572
      Top             =   3960
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   4
      Left            =   6360
      Picture         =   "Table.frx":19874
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   3
      Left            =   5760
      Picture         =   "Table.frx":1AB76
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   2
      Left            =   5160
      Picture         =   "Table.frx":1BE78
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   1
      Left            =   4560
      Picture         =   "Table.frx":1D17A
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image BM 
      Height          =   600
      Index           =   0
      Left            =   3960
      Picture         =   "Table.frx":1E47C
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   24
      Left            =   120
      Picture         =   "Table.frx":1F77E
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   23
      Left            =   120
      Picture         =   "Table.frx":20A80
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   22
      Left            =   120
      Picture         =   "Table.frx":223C2
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   21
      Left            =   120
      Picture         =   "Table.frx":23D04
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   20
      Left            =   120
      Picture         =   "Table.frx":25646
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   19
      Left            =   120
      Picture         =   "Table.frx":26F88
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   18
      Left            =   120
      Picture         =   "Table.frx":288CA
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   17
      Left            =   120
      Picture         =   "Table.frx":2A20C
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   16
      Left            =   120
      Picture         =   "Table.frx":2BB4E
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   15
      Left            =   120
      Picture         =   "Table.frx":2D490
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   14
      Left            =   120
      Picture         =   "Table.frx":2EDD2
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   13
      Left            =   120
      Picture         =   "Table.frx":30714
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   12
      Left            =   120
      Picture         =   "Table.frx":32056
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   11
      Left            =   120
      Picture         =   "Table.frx":33998
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   10
      Left            =   120
      Picture         =   "Table.frx":352DA
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   9
      Left            =   120
      Picture         =   "Table.frx":36C1C
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   8
      Left            =   120
      Picture         =   "Table.frx":3855E
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   7
      Left            =   120
      Picture         =   "Table.frx":39EA0
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   6
      Left            =   120
      Picture         =   "Table.frx":3B7E2
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   5
      Left            =   120
      Picture         =   "Table.frx":3D124
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   4
      Left            =   120
      Picture         =   "Table.frx":3EA66
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   3
      Left            =   120
      Picture         =   "Table.frx":403A8
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   2
      Left            =   120
      Picture         =   "Table.frx":41CEA
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   1
      Left            =   120
      Picture         =   "Table.frx":4362C
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Block 
      Height          =   600
      Index           =   0
      Left            =   120
      Picture         =   "Table.frx":44F6E
      Top             =   -2.45760e5
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu Lvl 
      Caption         =   "Levels"
      Begin VB.Menu l1 
         Caption         =   "Level 1"
      End
      Begin VB.Menu l2 
         Caption         =   "Level2"
      End
      Begin VB.Menu l3 
         Caption         =   "Level 3"
      End
   End
   Begin VB.Menu abt 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsEn(24) As Boolean
Dim Levels As Integer  ' Values 1 , 2 & 3

Private Sub abt_Click()
    HlpFrm.Show
End Sub

Private Sub BM_Click(Index As Integer)
If Levels = 3 Then
    Select Case (Index)
        Case 0:
            MoveLeft (Index)
            MoveDown (Index)
        Case 1, 2, 3:
            MoveDown (Index)
            MoveRight (Index)
            MoveLeft (Index)
        Case 4:
            MoveDown (Index)
            MoveRight (Index)
        Case 5, 10, 15:
            MoveUp (Index)
            MoveDown (Index)
            MoveLeft (Index)
        Case 20:
            MoveUp (Index)
            MoveLeft (Index)
        Case 9, 14, 19:
            MoveUp (Index)
            MoveDown (Index)
            MoveRight (Index)
        Case 24:
            MoveUp (Index)
            MoveRight (Index)
        Case 21, 22, 23:
            MoveUp (Index)
            MoveRight (Index)
            MoveLeft (Index)
        Case 6, 7, 8, 11, 12, 13, 16, 17, 18, 21, 22, 23:
            MoveUp (Index)
            MoveDown (Index)
            MoveRight (Index)
            MoveLeft (Index)
    End Select
ElseIf Levels = 2 Then
    Select Case (Index)
        Case 0:
                MoveDown (Index)
                MoveLeft (Index)
        Case 1, 2:
                MoveDown (Index)
                MoveLeft (Index)
                MoveRight (Index)
        Case 3:
                MoveDown (Index)
                MoveRight (Index)
        Case 4, 8:
                MoveUp (Index)
                MoveDown (Index)
                MoveLeft (Index)
        Case 12:
                MoveUp (Index)
                MoveLeft (Index)
        Case 5, 6, 9, 10:
                MoveUp (Index)
                MoveDown (Index)
                MoveLeft (Index)
                MoveRight (Index)
        Case 7, 11:
                MoveUp (Index)
                MoveDown (Index)
                MoveRight (Index)
        Case 13, 14:
                MoveUp (Index)
                MoveLeft (Index)
                MoveRight (Index)
        Case 15:
                MoveUp (Index)
                MoveRight (Index)
    End Select
Else
    Select Case (Index)
        Case 0:
                MoveLeft (Index)
                MoveDown (Index)
        Case 1:
                MoveLeft (Index)
                MoveDown (Index)
                MoveRight (Index)
        Case 2:
                MoveRight (Index)
                MoveDown (Index)
        Case 3:
                MoveUp (Index)
                MoveDown (Index)
                MoveLeft (Index)
        Case 4:
                MoveUp (Index)
                MoveDown (Index)
                MoveLeft (Index)
                MoveRight (Index)
        Case 5:
                MoveUp (Index)
                MoveDown (Index)
                MoveRight (Index)
        Case 6:
                MoveUp (Index)
                MoveLeft (Index)
        Case 7:
                MoveUp (Index)
                MoveLeft (Index)
                MoveRight (Index)
        Case 8:
                MoveUp (Index)
                MoveRight (Index)
    End Select
End If
        DidWin
End Sub
Function MoveUp(X As Integer)
        If (BM(X - (Levels + 2)).Picture = Block(24).Picture) Then
            BM(X - (Levels + 2)).Picture = BM(X).Picture
            BM(X).Picture = Block(24).Picture
        End If
End Function
Function MoveDown(X As Integer)
        If (BM(X + (Levels + 2)).Picture = Block(24).Picture) Then
            BM(X + (Levels + 2)).Picture = BM(X).Picture
            BM(X).Picture = Block(24).Picture
        End If
End Function
Function MoveLeft(X As Integer)
        If (BM(X + 1).Picture = Block(24).Picture) Then
            BM(X + 1).Picture = BM(X).Picture
            BM(X).Picture = Block(24).Picture
        End If
End Function
Function MoveRight(X As Integer)
      If (BM(X - 1).Picture = Block(24).Picture) Then
            BM(X - 1).Picture = BM(X).Picture
            BM(X).Picture = Block(24).Picture
      End If
End Function
Private Sub Command1_Click()
Dim T As Integer, C As Integer
C = 0
Order
For i = 0 To 24
        BM(i).Picture = Block(24).Picture
        IsEn(i) = False
Next i
If Levels = 3 Then
    While (C <= 24)
        Randomize (Second(Now))
        T = Int(25 * Rnd())
        If (IsEn(T) = False) Then
            IsEn(T) = True
            BM(C).Picture = Block(T).Picture
            C = C + 1
        End If
    Wend
ElseIf Levels = 2 Then
    For i = 16 To 24
        IsEn(i) = True
    Next i
    While (C <= 14)
        Randomize (Second(Now))
        T = Int(15 * Rnd())
        If (IsEn(T) = False) Then
            IsEn(T) = True
            BM(C).Picture = Block(T).Picture
            C = C + 1
        End If
    Wend
Else
    For i = 9 To 24
        IsEn(i) = True
    Next i
    While (C <= 7)
        Randomize (Second(Now))
        T = Int(8 * Rnd())
        If (IsEn(T) = False) Then
            IsEn(T) = True
            BM(C).Picture = Block(T).Picture
            C = C + 1
        End If
    Wend

End If
End Sub
Private Sub Form_Load()
Levels = 3
Call Command1_Click
l3.Checked = True
End Sub
Private Sub l1_Click()
Levels = 1
Call Command1_Click
l1.Checked = True
l2.Checked = False
l3.Checked = False
End Sub

Private Sub l2_Click()
Levels = 2
Call Command1_Click
l1.Checked = False
l2.Checked = True
l3.Checked = False
End Sub

Private Sub l3_Click()
Levels = 3
Call Command1_Click
l1.Checked = False
l2.Checked = False
l3.Checked = True
End Sub
Function Order()
    For k = 0 To 24
        BM(k).Left = -60000
        BM(k).Top = -60000
    Next k
    Dim C As Integer
    C = 0
    i = 120
    j = 120
    If Levels = 3 Then
        While C <= 24

                    BM(C).Top = i
                    BM(C).Left = j
                    j = j + 600
                    If j > 2520 Then
                        j = 120
                        i = i + 600
                    End If
            C = C + 1
        Wend
    ElseIf Levels = 2 Then
        While C <= 15

                    BM(C).Top = i
                    BM(C).Left = j
                    j = j + 600
                    If j > 1920 Then
                        j = 120
                        i = i + 600
                    End If
            C = C + 1
        Wend
    Else
            While C <= 8

                    BM(C).Top = i
                    BM(C).Left = j
                    j = j + 600
                    If j > 1320 Then
                        j = 120
                        i = i + 600
                    End If
            C = C + 1
            Wend
    End If

End Function
Function DidWin()
  Dim Flag As Boolean
   Flag = True
    If Levels = 1 Then
        For i = 0 To 7
            If BM(i).Picture <> Block(i).Picture Then
                Flag = False
            End If
        Next i
    ElseIf Levels = 2 Then
          For i = 0 To 14
            If BM(i).Picture <> Block(i).Picture Then
                Flag = False
            End If
        Next i
    Else
        For i = 0 To 23
            If BM(i).Picture <> Block(i).Picture Then
                Flag = False
            End If
        Next i
    End If
    If Flag = True Then
        Won.Show
        Call Command1_Click
    End If
End Function
