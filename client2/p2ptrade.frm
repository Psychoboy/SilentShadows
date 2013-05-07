VERSION 5.00
Begin VB.Form p2ptrade 
   BackColor       =   &H00000000&
   Caption         =   "Trade"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form3"
   ScaleHeight     =   2655
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LSTINV 
      Height          =   1230
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   2760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   7080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click an item you own, then click the item button for what slot you want to have it in"
      ForeColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4920
      Picture         =   "p2ptrade.frx":0000
      Top             =   1680
      Width           =   600
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4920
      Picture         =   "p2ptrade.frx":0F44
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4920
      Picture         =   "p2ptrade.frx":1E88
      Top             =   240
      Width           =   600
   End
   Begin VB.Image MyI3 
      Height          =   480
      Left            =   2400
      Picture         =   "p2ptrade.frx":2DCC
      Top             =   1680
      Width           =   600
   End
   Begin VB.Image MyI2 
      Height          =   480
      Left            =   2400
      Picture         =   "p2ptrade.frx":3D10
      Top             =   960
      Width           =   600
   End
   Begin VB.Image MyI1 
      Height          =   480
      Left            =   2400
      Picture         =   "p2ptrade.frx":4C54
      Top             =   240
      Width           =   600
   End
   Begin VB.Label YI3 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label YI2 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label YI1 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label OPI3 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label OPI2 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label OPI1 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   7200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   2640
   End
End
Attribute VB_Name = "p2ptrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MyI1_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
YI1.Caption = "1) " & Item(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)).name
MyI1.Tag = LSTINV.ListIndex + 1

If MyI1.Tag = MyI2.Tag Then
MyI2.Tag = 0
YI2.Caption = "2)"
End If

If MyI1.Tag = MyI3.Tag Then
MyI3.Tag = 0
YI3.Caption = "3)"
End If
Call SendData("TRI1" & SEP_CHAR & Val(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)) & SEP_CHAR & END_CHAR)
End Sub

Private Sub Image2_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Image3_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub MyI2_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
YI2.Caption = "2) " & Item(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)).name
MyI2.Tag = LSTINV.ListIndex + 1

If MyI2.Tag = MyI1.Tag Then
MyI1.Tag = 0
YI1.Caption = "1)"
End If

If MyI2.Tag = MyI3.Tag Then
MyI3.Tag = 0
YI3.Caption = "3)"
End If
Dim Packet As String
Call SendData("TRI2" & SEP_CHAR & MyIndex & SEP_CHAR & Val(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)) & SEP_CHAR & END_CHAR)
End Sub

Private Sub MyI3_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
YI3.Caption = "3) " & Item(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)).name
MyI3.Tag = LSTINV.ListIndex + 1

If MyI3.Tag = MyI2.Tag Then
MyI2.Tag = 0
YI2.Caption = "2)"
End If

If MyI3.Tag = MyI1.Tag Then
MyI1.Tag = 0
YI1.Caption = "1)"
End If
Call SendData("TRI3" & SEP_CHAR & Val(GetPlayerInvItemNum(MyIndex, LSTINV.ListIndex + 1)) & SEP_CHAR & END_CHAR)
End Sub
