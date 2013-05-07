VERSION 5.00
Begin VB.Form FRMPKLIST 
   BackColor       =   &H00000000&
   Caption         =   "PK Log... 'n stuff"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   240
      Picture         =   "FRMPKLIST.frx":0000
      Top             =   5040
      Width           =   3105
   End
End
Attribute VB_Name = "FRMPKLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Me.Hide
End Sub
