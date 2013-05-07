VERSION 5.00
Begin VB.Form seesign 
   BackColor       =   &H00000000&
   Caption         =   "Book"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "seesign.frx":0000
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   480
      Picture         =   "seesign.frx":0006
      Top             =   3120
      Width           =   3105
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "seesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
Unload Me
End Sub

