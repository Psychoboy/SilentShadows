VERSION 5.00
Begin VB.Form deathedit 
   Caption         =   "Death Text"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "-------------->"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Sound?"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Message On Death : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "deathedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Deathsay = Text1.Text
SoundFileD = Text2.Text
Unload Me
End Sub

Private Sub File1_Click()
Text2.Text = File1.filename
End Sub

Private Sub File1_DblClick()

Call PlaySound("\sfx\" & File1.filename)
End Sub

Private Sub Form_Load()
'File1.Path = App.Path & "\sfx\"
End Sub

