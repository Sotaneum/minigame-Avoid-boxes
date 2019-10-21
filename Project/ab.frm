VERSION 5.00
Begin VB.Form ab 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "About"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "ab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "닫기"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "홈페이지 접속"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "대충만든겁니다. 그냥 Yes 누르고 상자 피하시면됩니다. 제작자 : Sotaneum"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4290
   End
End
Attribute VB_Name = "ab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "explorer.exe http://duration.digimoon.net/"
End

End Sub

Private Sub Command2_Click()
ab.Hide

End Sub
