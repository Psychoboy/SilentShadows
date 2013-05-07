VERSION 5.00
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   6855
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   2400
      TabIndex        =   36
      Text            =   "0"
      Top             =   5760
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1440
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2520
      TabIndex        =   23
      Top             =   3240
      Width           =   2415
   End
   Begin VB.HScrollBar ScrMPUSED 
      Height          =   375
      Left            =   1200
      Max             =   1024
      Min             =   1
      TabIndex        =   21
      Top             =   1920
      Value           =   1
      Width           =   3255
   End
   Begin VB.HScrollBar scrlLevelReq 
      Height          =   375
      Left            =   960
      Max             =   255
      Min             =   1
      TabIndex        =   18
      Top             =   1080
      Value           =   1
      Width           =   3495
   End
   Begin VB.ComboBox cmbClassReq 
      Height          =   390
      ItemData        =   "frmSpellEditor.frx":0000
      Left            =   120
      List            =   "frmSpellEditor.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   600
      Width           =   4815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   6240
      Width           =   2295
   End
   Begin VB.ComboBox cmbType 
      Height          =   390
      ItemData        =   "frmSpellEditor.frx":0004
      Left            =   120
      List            =   "frmSpellEditor.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame warpframe 
      Caption         =   "Crafting"
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   4815
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Crafted"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Item1"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label asdf 
         Caption         =   "Item2"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame fraGiveItem 
      Caption         =   "Give Item"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlItemValue 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   375
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   10
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblItemValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Value"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblItemNum 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   375
         Left            =   1320
         Max             =   1024
         TabIndex        =   3
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Gfxspell 
      Caption         =   "GFX for the spell :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Sound Effect"
      Height          =   615
      Left            =   3240
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label MPUSEDLBL 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "MP Used"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLevelReq 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Level"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbType_Click()
    If cmbType.ListIndex <> SPELL_TYPE_GIVEITEM Then
        fraVitals.Visible = True
        fraGiveItem.Visible = False
        warpframe.Visible = False
    Else
        fraVitals.Visible = False
        fraGiveItem.Visible = True
        warpframe.Visible = False
    End If
    If cmbType.ListIndex = SPELL_TYPE_CRAFT Then
        fraVitals.Visible = False
        fraGiveItem.Visible = False
        warpframe.Visible = True
    End If
        
End Sub

Private Sub File1_Click()
Text1.Text = File1.filename
End Sub

Private Sub File1_DblClick()
Call PlaySound("\sfx\" & File1.filename)
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\sfx\"
End Sub

Private Sub HScroll1_Change()
Label10.Caption = HScroll1.value & ". " & Item(HScroll1.value).name
End Sub

Private Sub HScroll2_Change()
Label11.Caption = HScroll2.value & ". " & Item(HScroll2.value).name
End Sub

Private Sub HScroll3_Change()
Label12.Caption = HScroll3.value & ". " & Item(HScroll3.value).name
End Sub

Private Sub scrlItemNum_Change()
    fraGiveItem.Caption = "Give Item " & ". " & Trim(Item(scrlItemNum.value).name)
    lblItemNum.Caption = STR(scrlItemNum.value)
End Sub

Private Sub scrlItemValue_Change()
    lblItemValue.Caption = STR(scrlItemValue.value)
End Sub

Private Sub scrlLevelReq_Change()
    lblLevelReq.Caption = STR(scrlLevelReq.value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub ScrMPUSED_Change()
MPUSEDLBL.Caption = STR(ScrMPUSED.value)
End Sub
