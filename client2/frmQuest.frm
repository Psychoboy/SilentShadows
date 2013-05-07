VERSION 5.00
Begin VB.Form frmQuest 
   Caption         =   "Quest Editor"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   5415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   720
      Max             =   1000
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   720
      Max             =   1000
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "NAME"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "NAME"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Message :"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Given Item :"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Required Item :"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Quest1 = HScroll1.value
Quest2 = Text1.Text
Quest3 = HScroll2.value

Unload Me
End Sub

Private Sub HScroll1_Change()
Label5.Caption = HScroll1.value
Label6.Caption = Item(HScroll1.value).name
End Sub

Private Sub HScroll2_Change()
Label4.Caption = HScroll2.value
Label7.Caption = Item(HScroll2.value).name
End Sub
