VERSION 5.00
Begin VB.Form frmItemInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Info"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label magimod 
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label defmod 
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label strmod 
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Class 
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label strdeflvl 
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Durability 
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "MAGI Mod"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "DEF Mod"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "STR Mod"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Class"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label strdeflvl2 
      Caption         =   "STR/DEF/LVL"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Durability"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmItemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmItemInfo.Hide
End Sub
