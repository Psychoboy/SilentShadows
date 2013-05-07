VERSION 5.00
Begin VB.Form GFX 
   Caption         =   "Form1"
   ClientHeight    =   540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.Image spells 
      Height          =   37500
      Left            =   7080
      Picture         =   "GFX.frx":0000
      Top             =   7920
      Width           =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image sprites 
      Height          =   1.50000e5
      Left            =   7560
      Picture         =   "GFX.frx":2BF244
      Top             =   6120
      Width           =   5760
   End
   Begin VB.Image tiles 
      Height          =   1.12500e5
      Left            =   7200
      Picture         =   "GFX.frx":DBBA88
      Top             =   4200
      Width           =   3360
   End
   Begin VB.Image items 
      Height          =   2.25000e5
      Left            =   10800
      Picture         =   "GFX.frx":128A24C
      Top             =   3480
      Width           =   480
   End
End
Attribute VB_Name = "GFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()

End Sub

