VERSION 5.00
Begin VB.Form Healsound 
   Caption         =   "Heal"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   1680
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "------------->"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Sound?"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Healsound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SoundFileH = Text1.Text
Unload Me
End Sub

Private Sub File1_DblClick()
Text1.Text = File1.filename
End Sub

Private Sub Form_Load()
'File1.Path = App.Path & "\sfx\"
End Sub
