VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton enable 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   0
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   2160
      Width           =   480
   End
   Begin VB.PictureBox picPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Interval = 50
End Sub

Private Sub Timer1_Timer()
Call WeaponBltItem
End Sub
