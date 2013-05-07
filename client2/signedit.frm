VERSION 5.00
Begin VB.Form signedit 
   Caption         =   "Sign Editor"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   840
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Sound?"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "signedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SignTitle = Text1.Text
SignText = Text2.Text
SoundFileS = Text3.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub File1_Click()
Text3.Text = File1.filename
End Sub

Private Sub File1_DblClick()

Call PlaySound("\sfx\" & File1.filename)
End Sub

Private Sub Form_Load()
'File1.Path = App.Path & "\sfx\"
End Sub
