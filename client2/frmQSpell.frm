VERSION 5.00
Begin VB.Form frmQSpell 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Spell Set Up"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   840
      Left            =   360
      Picture         =   "frmQSpell.frx":0000
      Top             =   2640
      Width           =   3105
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   120
      Picture         =   "frmQSpell.frx":3345
      Top             =   1800
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   120
      Picture         =   "frmQSpell.frx":6238
      Top             =   960
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "frmQSpell.frx":914C
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmQSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call SendData("spells" & SEP_CHAR & END_CHAR)

End Sub

Private Sub Image4_Click()
Me.Hide
End Sub

Private Sub List1_Click()
QUICKSPELL1 = List1.ListIndex
End Sub

Private Sub List2_Click()
QUICKSPELL2 = List2.ListIndex
End Sub

Private Sub List3_Click()
QUICKSPELL3 = List3.ListIndex
End Sub
