VERSION 5.00
Begin VB.Form SFXtile 
   Caption         =   "SFXTile"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form2"
   ScaleHeight     =   2145
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox Check3 
         Caption         =   "Shallow"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Real Deep"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Deep Tile"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "SFXtile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SoundFile = Text1.text
Deepness = 0
If Check1.value = 1 Then Deepness = 9998
If Check2.value = 1 Then Deepness = 9999
If Check3.value = 1 Then Deepness = 9997

Unload Me
End Sub

Private Sub File1_Click()
Text1.text = File1.filename
End Sub

Private Sub File1_DblClick()
Call PlaySound("\sfx\" & File1.filename)
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\sfx\"
End Sub
