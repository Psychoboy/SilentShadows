VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
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
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtND 
      Height          =   390
      Left            =   3840
      TabIndex        =   72
      Text            =   "0"
      Top             =   2160
      Width           =   555
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   255
      Left            =   2640
      TabIndex        =   67
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox Picpic2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   62
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItems2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   61
      Top             =   7200
      Width           =   480
   End
   Begin VB.TextBox Text4 
      Height          =   390
      Left            =   1140
      TabIndex        =   59
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   720
      TabIndex        =   57
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SFXtest"
      Height          =   390
      Left            =   120
      TabIndex        =   34
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox cmbClass 
      Height          =   390
      ItemData        =   "frmItemEditor.frx":0000
      Left            =   120
      List            =   "frmItemEditor.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Timer tmrPic 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   7200
      Width           =   480
   End
   Begin VB.PictureBox picPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   600
      Width           =   480
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   375
      Left            =   960
      Max             =   600
      TabIndex        =   17
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox cmbType 
      Height          =   390
      ItemData        =   "frmItemEditor.frx":00E9
      Left            =   120
      List            =   "frmItemEditor.frx":0129
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar SPDMODER 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   76
         Top             =   1800
         Width           =   2895
      End
      Begin VB.HScrollBar scrlrange 
         Height          =   255
         Left            =   1320
         Max             =   20
         TabIndex        =   73
         Top             =   1200
         Value           =   1
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.HScrollBar MAGIMODER 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   39
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar DEFMODER 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   38
         Top             =   2280
         Width           =   2895
      End
      Begin VB.HScrollBar STRMODER 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   37
         Top             =   2040
         Width           =   2895
      End
      Begin VB.HScrollBar scrlsound 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   31
         Top             =   960
         Value           =   1
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   1320
         Max             =   1000
         TabIndex        =   8
         Top             =   720
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   6
         Top             =   480
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label26 
         Caption         =   "Modifiers"
         Height          =   255
         Left            =   2160
         TabIndex        =   79
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblspeed 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4200
         TabIndex        =   78
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblrange 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   77
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Range"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Magi"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Defence"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Strength mod"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4200
         TabIndex        =   42
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4200
         TabIndex        =   41
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4200
         TabIndex        =   40
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblsound 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Sound"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Strength"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Durability"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame ringammy 
      Caption         =   "Book"
      Height          =   2895
      Left            =   120
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   1095
         Left            =   240
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.Frame SChanger 
      Caption         =   "Sprite change"
      Height          =   2895
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   4815
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   28
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   2895
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlSpell 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   22
         Top             =   840
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Num"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSpell 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblSpellName 
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   375
         Left            =   1320
         Max             =   255
         TabIndex        =   12
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scrolls"
      Height          =   2895
      Left            =   120
      TabIndex        =   68
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar HScroll5 
         Height          =   375
         Left            =   240
         Max             =   255
         TabIndex        =   69
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   495
         Left            =   2760
         TabIndex        =   70
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FrmEgg 
      Caption         =   "Pet Eggs"
      Height          =   2895
      Left            =   120
      TabIndex        =   63
      Top             =   3240
      Width           =   4815
      Begin VB.HScrollBar HScroll6 
         Height          =   495
         Left            =   240
         Max             =   2000
         Min             =   10
         TabIndex        =   64
         Top             =   1200
         Value           =   10
         Width           =   3255
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Pet's Sprite"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label egglab 
         Caption         =   "0"
         Height          =   375
         Left            =   3720
         TabIndex        =   65
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame FRMmaskedit 
      Caption         =   "Mask / Fringe Editor"
      Height          =   2895
      Left            =   120
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox fringemask 
         Caption         =   "Fringe instead of Mask"
         Height          =   375
         Left            =   840
         TabIndex        =   56
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   2040
         TabIndex        =   54
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Tile Number"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame randframe 
      Caption         =   "RandItem"
      Height          =   2895
      Left            =   120
      TabIndex        =   46
      Top             =   3240
      Width           =   4815
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   480
         Max             =   1000
         TabIndex        =   49
         Top             =   480
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   480
         Max             =   1000
         TabIndex        =   48
         Top             =   1080
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   480
         Max             =   1000
         TabIndex        =   47
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   495
         Left            =   2160
         TabIndex        =   52
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Label16"
         Height          =   495
         Left            =   2160
         TabIndex        =   51
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Label16"
         Height          =   495
         Left            =   2160
         TabIndex        =   50
         Top             =   1680
         Width           =   2415
      End
   End
   Begin VB.Label lblND 
      Caption         =   "ND"
      Height          =   315
      Left            =   3300
      TabIndex        =   71
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "Sell Value"
      Height          =   375
      Left            =   0
      TabIndex        =   60
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "Sprite"
      Height          =   375
      Left            =   0
      TabIndex        =   58
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblPic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Pic"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
    Else
        fraEquipment.Visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
    Label15.Visible = True
    scrlsound.Visible = True
    lblsound.Visible = True
    Label24.Visible = True
    scrlRange.Visible = True
    lblRange.Visible = True
    Else
    Label15.Visible = False
    scrlsound.Visible = False
    lblsound.Visible = False
    Label24.Visible = False
    scrlRange.Visible = False
    lblRange.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_MASKEDIT) Then
        FRMmaskedit.Visible = True
    Else
        FRMmaskedit.Visible = False
    End If
    If (cmbType.ListIndex = ITEM_TYPE_SCHANGE) Then
        SChanger.Visible = True
    Else
        SChanger.Visible = False
    End If
    If (cmbType.ListIndex = ITEM_TYPE_BOOK) Then
        ringammy.Visible = True
    Else
        ringammy.Visible = False
    End If
    
     If (cmbType.ListIndex = ITEM_TYPE_RANDITEM) Then
        randframe.Visible = True
    Else
        randframe.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_PETEGG) Then
        FrmEgg.Visible = True
    Else
        FrmEgg.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_BARBER) Then
        SChanger.Visible = True
    Else
        SChanger.Visible = False
    End If
    
    
    
End Sub

Private Sub Command1_Click()
Call PlaySound("\sfx\weap" & frmItemEditor.scrlsound.value & ".wav")
End Sub

Private Sub DEFMODER_Change()
Label10.Caption = STR(DEFMODER.value)
End Sub

Private Sub Form_Load()
frmItemEditor.cmbClass.Clear
Dim i As Integer
frmItemEditor.cmbClass.AddItem "None"
        For i = 0 To Max_Classes
            frmItemEditor.cmbClass.AddItem Trim(Class(i).name)
        Next i
            
        frmItemEditor.cmbClass.ListIndex = 0
End Sub

Private Sub HScroll1_Change()
Label8.Caption = HScroll1.value
End Sub

Private Sub HScroll2_Change()
Label16.Caption = Item(HScroll2.value).name
End Sub

Private Sub HScroll3_Change()
Label17.Caption = Item(HScroll3.value).name
End Sub

Private Sub HScroll4_Change()
Label18.Caption = Item(HScroll4.value).name
End Sub

Private Sub HScroll5_Change()
Label20.Caption = Spell(HScroll5.value).name
End Sub

Private Sub HScroll6_Change()
egglab.Caption = HScroll6.value
End Sub

Private Sub HScroll7_Change()
Text3.Text = HScroll7.value
End Sub

Private Sub MAGIMODER_Change()
Label11.Caption = STR(MAGIMODER.value)
End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = STR(scrlPic.value)
End Sub

Private Sub scrlRange_Change()

lblRange.Caption = STR(scrlRange.value)
End Sub

Private Sub scrlsound_Change()
lblsound.Caption = STR(scrlsound.value)
End Sub

Private Sub SPDMODER_Change()
    lblSPEED.Caption = STR(SPDMODER.value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.value)
End Sub

Private Sub scrlVitalAdd_Change()
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = STR(scrlDurability.value)
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim(Spell(scrlSpell.value).name)
    lblSpell.Caption = STR(scrlSpell.value)
End Sub

Private Sub STRMODER_Change()
Label9.Caption = STR(STRMODER.value)
End Sub

Private Sub Text3_Change()
HScroll7.value = Text3.Text
End Sub

Private Sub tmrPic_Timer()
    Call ItemEditorBltItem
End Sub


