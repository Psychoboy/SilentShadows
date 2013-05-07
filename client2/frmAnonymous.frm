VERSION 5.00
Begin VB.Form frmAnonymous 
   Caption         =   "Be Anonymous"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "NO"
      Height          =   495
      Left            =   2580
      TabIndex        =   2
      Top             =   780
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "YES"
      Height          =   495
      Left            =   540
      TabIndex        =   1
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAnonymous.frx":0000
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "frmAnonymous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("anonymous" & SEP_CHAR & END_CHAR)
frmAnonymous.Hide
End Sub

Private Sub Command2_Click()
frmAnonymous.Hide
End Sub
