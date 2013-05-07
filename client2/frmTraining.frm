VERSION 5.00
Begin VB.Form frmTraining 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Shadows (Training)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStat 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   405
      ItemData        =   "frmTraining.frx":0000
      Left            =   2340
      List            =   "frmTraining.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   0
      Picture         =   "frmTraining.frx":0059
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image4 
      Height          =   840
      Left            =   3120
      Picture         =   "frmTraining.frx":95BF
      Top             =   3360
      Width           =   3105
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   3120
      Picture         =   "frmTraining.frx":D5BA
      Top             =   2640
      Width           =   3105
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   3000
      Picture         =   "frmTraining.frx":11045
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What stat would you like to train?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
End
Attribute VB_Name = "frmTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cmbStat.ListIndex = 0
End Sub

Private Sub Image3_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendData("usestatpoint" & SEP_CHAR & cmbStat.ListIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Image4_Click()
sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
Unload Me
End Sub
