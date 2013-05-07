VERSION 5.00
Begin VB.Form idiot 
   BorderStyle     =   0  'None
   Caption         =   "Error"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form2"
   Picture         =   "idiot.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4920
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "idiot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
