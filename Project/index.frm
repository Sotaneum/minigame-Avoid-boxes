VERSION 5.00
Begin VB.Form index 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Mini Game for Sotaneum"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2895
   Icon            =   "index.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame gom 
      Caption         =   "Game Over"
      Height          =   1815
      Left            =   360
      TabIndex        =   78
      Top             =   30000
      Width           =   2175
      Begin VB.CommandButton n 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   30000
         Width           =   1575
      End
      Begin VB.CommandButton y 
         Caption         =   "Yes"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   30000
         Width           =   1575
      End
      Begin VB.Label er 
         AutoSize        =   -1  'True
         Caption         =   "다시시도?"
         Height          =   180
         Left            =   120
         TabIndex        =   79
         Top             =   30000
         Width           =   810
      End
   End
   Begin VB.Timer Game_over 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton ab3 
      Caption         =   "About"
      Height          =   375
      Left            =   480
      TabIndex        =   77
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton no 
      Caption         =   "No"
      Height          =   375
      Left            =   1560
      TabIndex        =   75
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton yes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   480
      TabIndex        =   74
      Top             =   1800
      Width           =   735
   End
   Begin VB.Timer Main_Game_Start 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   0
   End
   Begin VB.Label sh 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Sotaneum HomePage"
      Height          =   180
      Left            =   480
      TabIndex        =   76
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Label start 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Game Start"
      Height          =   180
      Left            =   480
      TabIndex        =   73
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   70
      Left            =   960
      TabIndex        =   72
      Top             =   5280
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   69
      Left            =   960
      TabIndex        =   71
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   68
      Left            =   960
      TabIndex        =   70
      Top             =   5040
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   67
      Left            =   1200
      TabIndex        =   69
      Top             =   5520
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   66
      Left            =   960
      TabIndex        =   68
      Top             =   5040
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   65
      Left            =   1080
      TabIndex        =   67
      Top             =   5280
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   64
      Left            =   1080
      TabIndex        =   66
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   63
      Left            =   1080
      TabIndex        =   65
      Top             =   5040
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   62
      Left            =   1080
      TabIndex        =   64
      Top             =   5040
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   61
      Left            =   1200
      TabIndex        =   63
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   60
      Left            =   1200
      TabIndex        =   62
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   59
      Left            =   1200
      TabIndex        =   61
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   58
      Left            =   1440
      TabIndex        =   60
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   57
      Left            =   1440
      TabIndex        =   59
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   56
      Left            =   1440
      TabIndex        =   58
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   55
      Left            =   960
      TabIndex        =   57
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   54
      Left            =   960
      TabIndex        =   56
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   53
      Left            =   1080
      TabIndex        =   55
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   52
      Left            =   960
      TabIndex        =   54
      Top             =   5640
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   51
      Left            =   1200
      TabIndex        =   53
      Top             =   6240
      Width           =   150
   End
   Begin VB.Label stage 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "│Score : 0│Lv.1│"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   1650
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   50
      Left            =   1440
      TabIndex        =   51
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   49
      Left            =   1440
      TabIndex        =   50
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   48
      Left            =   1440
      TabIndex        =   49
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   47
      Left            =   1440
      TabIndex        =   48
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   46
      Left            =   1440
      TabIndex        =   47
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   45
      Left            =   1440
      TabIndex        =   46
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   44
      Left            =   1440
      TabIndex        =   45
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   43
      Left            =   1440
      TabIndex        =   44
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   42
      Left            =   1440
      TabIndex        =   43
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   41
      Left            =   1440
      TabIndex        =   42
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   40
      Left            =   1440
      TabIndex        =   41
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   39
      Left            =   1440
      TabIndex        =   40
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   38
      Left            =   1440
      TabIndex        =   39
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   37
      Left            =   1440
      TabIndex        =   38
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   36
      Left            =   1440
      TabIndex        =   37
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   35
      Left            =   1440
      TabIndex        =   36
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   34
      Left            =   1440
      TabIndex        =   35
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   33
      Left            =   1440
      TabIndex        =   34
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   32
      Left            =   1440
      TabIndex        =   33
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   31
      Left            =   1440
      TabIndex        =   32
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   30
      Left            =   1440
      TabIndex        =   31
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   29
      Left            =   1440
      TabIndex        =   30
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   28
      Left            =   1440
      TabIndex        =   29
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   27
      Left            =   1440
      TabIndex        =   28
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   26
      Left            =   1440
      TabIndex        =   27
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   25
      Left            =   1440
      TabIndex        =   26
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   24
      Left            =   1440
      TabIndex        =   25
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   23
      Left            =   1440
      TabIndex        =   24
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   22
      Left            =   1440
      TabIndex        =   23
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   21
      Left            =   1440
      TabIndex        =   22
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   20
      Left            =   1440
      TabIndex        =   21
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   19
      Left            =   1440
      TabIndex        =   20
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   18
      Left            =   1440
      TabIndex        =   19
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   17
      Left            =   1440
      TabIndex        =   18
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   16
      Left            =   1440
      TabIndex        =   17
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   15
      Left            =   1440
      TabIndex        =   16
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   14
      Left            =   1440
      TabIndex        =   15
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   13
      Left            =   1440
      TabIndex        =   14
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   12
      Left            =   1440
      TabIndex        =   13
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   11
      Left            =   1440
      TabIndex        =   12
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   10
      Left            =   1440
      TabIndex        =   11
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   8
      Left            =   1440
      TabIndex        =   9
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   7
      Left            =   1440
      TabIndex        =   8
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   6
      Left            =   1440
      TabIndex        =   7
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label Box 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "□"
      Height          =   180
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   6000
      Width           =   150
   End
   Begin VB.Label User 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "■"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   3840
      Width           =   150
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Score, Lv As Long

Private Sub ab3_Click()
ab.Show

End Sub

Private Sub Form_Load()
For i = 0 To 70 Step 1
Box(i).Top = Box(i).Top + 1 + Rnd() * i
If Box(i).Top > index.Height Then
    Box(i).Top = -600
    Box(i).Left = Rnd * i * 300
    
    If Box(i).Left > index.Width Then
        Box(i).Left = Rnd * i * 300
    End If

End If
Next i

End Sub

Private Sub Game_over_Timer()
gom.Top = 1200
gom.Left = 360
er.Top = 360
er.Left = 120
y.Top = 720
y.Left = 240
n.Top = 1200
n.Left = 240
ab3.Enabled = False
n.Enabled = True
no.Enabled = False
y.Enabled = True
yes.Enabled = False

End Sub

Private Sub n_Click()
End

End Sub

Private Sub sh_Click()
Shell "explorer.exe http://duration.digimoon.net/"
End

End Sub

Private Sub no_Click()
End

End Sub

Private Sub y_Click()
For i = 0 To 70 Step 1
Box(i).Top = 30000
Next i
User.Top = 3840
User.Left = 1320
Game_over.Enabled = False
gom.Top = 30000
gom.Left = 30000
er.Top = 30000
er.Left = 30000
y.Top = 30000
y.Left = 30000
n.Top = 30000
n.Left = 30000
Score = 0
Lv = 1
ab3.Enabled = False
n.Enabled = False
no.Enabled = False
y.Enabled = False
yes.Enabled = False
Main_Game_Start.Enabled = True
End Sub

Private Sub yes_Click()
For i = 0 To 70 Step 1
Box(i).Top = Box(i).Top + 1 + Rnd() * i
If Box(i).Top > index.Height Then
    Box(i).Top = -600
    Box(i).Left = Rnd * i * 300
    
    If Box(i).Left > index.Width Then
        Box(i).Left = Rnd * i * 300
    End If

End If
Next i
start.Top = 30000
yes.Top = 30000
no.Top = 30000
sh.Top = 30000
ab3.Top = 30000
Main_Game_Start.Enabled = True
ab3.Enabled = False
n.Enabled = False
no.Enabled = False
y.Enabled = False
yes.Enabled = False

End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DDown, ADown, WDown, SDown
Dim a, b As Long

DDown = GetAsyncKeyState(vbKeyD)
ADown = GetAsyncKeyState(vbKeyA)
WDown = GetAsyncKeyState(vbKeyW)
SDown = GetAsyncKeyState(vbKeyS)

a = User.Left
b = User.Top

If DDown Then

    User.Left = a + 100
    If User.Left > index.Width Then
        User.Left = 0
    End If
       
ElseIf ADown Then

    User.Left = a - 100
    If User.Left < 0 Then
        User.Left = index.Width
    End If
    
ElseIf WDown Then

    User.Top = b - 110
    If User.Top < 0 Then
        User.Top = 0
    End If
    
ElseIf SDown Then

    User.Top = b + 110
    If User.Top > index.Height Then
        User.Top = index.Height
    End If
End If

End Sub

Private Sub Main_Game_Start_Timer()
ab3.Enabled = False
n.Enabled = False
no.Enabled = False
y.Enabled = False
yes.Enabled = False
If Score = "" Then
Score = 0
End If

If Score <= 1000 Then
Lv = 1

For i = 0 To 40 Step 1
Box(i).Top = Box(i).Top + 1 + Rnd() * i
If Box(i).Top > index.Height Then
    Box(i).Top = -600
    Box(i).Left = Rnd * i * 300
    Score = Score + 10
    
    If Box(i).Left > index.Width Then
        Box(i).Left = Rnd * i * 300
    End If

End If

If Box(i).Top > User.Top - 110 Then
    If Box(i).Top < User.Top + 110 Then
        If Box(i).Left > User.Left - 95 Then
            If Box(i).Left < User.Left + 95 Then
                Game_over.Enabled = True
                Main_Game_Start.Enabled = False
            End If
        End If
    End If
End If
Next i

Else
Lv = 2
For i = 0 To 70 Step 1
Box(i).Top = Box(i).Top + 1 + Rnd() * i
If Box(i).Top > index.Height Then
    Box(i).Top = -600
    Box(i).Left = Rnd * i * 300
    Score = Score + 10
    
    If Box(i).Left > index.Width Then
        Box(i).Left = Rnd * i * 300
    End If

End If

If Box(i).Top > User.Top - 110 Then
    If Box(i).Top < User.Top + 110 Then
        If Box(i).Left > User.Left - 95 Then
            If Box(i).Left < User.Left + 95 Then
                Game_over.Enabled = True
                Main_Game_Start.Enabled = False
            End If
        End If
    End If
End If
Next i
End If

stage.Caption = "│Score : " & Score & "│Lv." & Lv & "│"

End Sub

