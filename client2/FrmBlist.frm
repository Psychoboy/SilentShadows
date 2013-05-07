VERSION 5.00
Begin VB.Form FrmBlist 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Buddy"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2295
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   240
      Max             =   10
      Min             =   1
      TabIndex        =   1
      Top             =   960
      Value           =   1
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add a Buddy!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   1560
      Picture         =   "FrmBlist.frx":0000
      Top             =   1560
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "FrmBlist.frx":2D3D
      Top             =   1560
      Width           =   570
   End
End
Attribute VB_Name = "FrmBlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
Label1.Caption = HScroll1.value
End Sub

Private Sub Image1_Click()
BuDDy(HScroll1.value) = Text1.Text
Call SaveBuddyX
Me.Hide
Call updateBList
End Sub

Private Sub Image2_Click()
Me.Hide
End Sub
