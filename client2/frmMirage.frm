VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMirage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Silent Shadows"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   11775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMirage.frx":0000
   ScaleHeight     =   576
   ScaleMode       =   0  'User
   ScaleWidth      =   802.043
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7485
      Left            =   7980
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Map ALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   31
         Top             =   6840
         Width           =   2055
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   3600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAttribs 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   5640
         Width           =   2055
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   5040
         Width           =   2055
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   3375
         Left            =   3480
         Max             =   1000
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picSelect 
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
         Left            =   3120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   3600
         Width           =   480
      End
      Begin VB.Frame fraLayers 
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3915
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
         Begin VB.OptionButton optF2Anim 
            Caption         =   "Animation"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optFringe2 
            Caption         =   "Fringe 2"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   2040
            Width           =   1215
         End
         Begin VB.OptionButton optFAnim 
            Caption         =   "Animation"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton optM2Anim 
            Caption         =   "Animation"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optMask2 
            Caption         =   "Mask2"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Fill"
            Height          =   270
            Left            =   120
            TabIndex        =   27
            Top             =   3000
            Width           =   1215
         End
         Begin VB.OptionButton optGround 
            Caption         =   "Ground"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optMask 
            Caption         =   "Mask"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optAnim 
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optFringe 
            Caption         =   "Fringe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.Frame fraAttribs 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3915
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton optMining 
            Caption         =   "Mining"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   3300
            Width           =   855
         End
         Begin VB.OptionButton optFishing 
            Caption         =   "Fishing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   70
            Top             =   3060
            Width           =   855
         End
         Begin VB.OptionButton optArena 
            Caption         =   "Arena"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   2820
            Width           =   1215
         End
         Begin VB.OptionButton optQuest 
            Caption         =   "Quest"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   2580
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Fill"
            Height          =   270
            Left            =   120
            TabIndex        =   30
            Top             =   3600
            Width           =   1215
         End
         Begin VB.OptionButton SFXopt 
            Caption         =   "SFX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   29
            Top             =   2340
            Width           =   1215
         End
         Begin VB.OptionButton deathopt 
            Caption         =   "Death"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   2100
            Width           =   1215
         End
         Begin VB.OptionButton fheal 
            Caption         =   "Full Heal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   1860
            Width           =   1215
         End
         Begin VB.OptionButton Optsign 
            Caption         =   "Sign"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   1620
            Width           =   1215
         End
         Begin VB.OptionButton optKeyOpen 
            Caption         =   "Key Open"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   1380
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            Caption         =   "Blocked"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optWarp 
            Caption         =   "Warp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   420
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   14
            Top             =   660
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            Caption         =   "Npc Avoid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton optKey 
            Caption         =   "Key"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   1140
            Width           =   1215
         End
      End
      Begin VB.PictureBox picBack 
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
         Height          =   3360
         Left            =   120
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   3
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
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
            Height          =   960
            Left            =   0
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   4
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.Label infolbl 
         Caption         =   "info"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   4440
         Width           =   375
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00800000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   8040
      Width           =   915
   End
   Begin RichTextLib.RichTextBox txtActions 
      Height          =   1215
      Left            =   8640
      TabIndex        =   72
      Top             =   5580
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":5408
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1605
      Left            =   240
      TabIndex        =   67
      Top             =   7080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2831
      _Version        =   393216
      TabOrientation  =   2
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "frmMirage.frx":548A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtChat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Global"
      TabPicture(1)   =   "frmMirage.frx":54A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCombat"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox txtCombat 
         Height          =   1635
         Left            =   -74700
         TabIndex        =   69
         Top             =   0
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2884
         _Version        =   393217
         BackColor       =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMirage.frx":54C2
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   1575
         Left            =   360
         TabIndex        =   68
         Top             =   0
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   2778
         _Version        =   393217
         BackColor       =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMirage.frx":5544
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   9720
      TabIndex        =   66
      Top             =   7080
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox MyTextBox 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   6600
      Width           =   8055
   End
   Begin VB.PictureBox picPic2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10380
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   61
      Top             =   4920
      Width           =   480
   End
   Begin VB.PictureBox picPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9660
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   60
      Top             =   4920
      Width           =   480
   End
   Begin VB.PictureBox picPic3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   8940
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   59
      Top             =   4920
      Width           =   480
   End
   Begin VB.PictureBox ITEM2 
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
      Left            =   9645
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   56
      Tag             =   "0"
      Top             =   4095
      Width           =   480
   End
   Begin VB.PictureBox ITEM1 
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
      Left            =   8925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   55
      Tag             =   "0"
      Top             =   4095
      Width           =   480
   End
   Begin VB.PictureBox ITEM3 
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
      Left            =   10365
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   53
      Tag             =   "0"
      Top             =   4095
      Width           =   480
   End
   Begin VB.PictureBox picRSM 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   240
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1200
         TabIndex        =   47
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ListBox BList 
         Height          =   1740
         ItemData        =   "frmMirage.frx":55C4
         Left            =   0
         List            =   "frmMirage.frx":55CB
         TabIndex        =   46
         Top             =   600
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2535
         Left            =   1200
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4471
         _Version        =   393217
         TextRTF         =   $"frmMirage.frx":55D6
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   0
         Picture         =   "frmMirage.frx":5658
         Top             =   2880
         Width           =   600
      End
      Begin VB.Image Image15 
         Height          =   480
         Left            =   600
         Picture         =   "frmMirage.frx":659C
         Top             =   2880
         Width           =   600
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   0
         Picture         =   "frmMirage.frx":74E0
         Top             =   2400
         Width           =   600
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   600
         Picture         =   "frmMirage.frx":8424
         Top             =   2400
         Width           =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         X1              =   0
         X2              =   232
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Friends List"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SS Messenger"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame fmeparty 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   10200
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      Begin VB.ListBox List1g 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Height          =   1230
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Party"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Ambiance3 
      Left            =   11760
      Top             =   960
   End
   Begin VB.Timer Ambiance2 
      Left            =   11760
      Top             =   480
   End
   Begin VB.Timer Ambiance1 
      Left            =   11760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   120
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   12120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   28
      Top             =   8520
      Width           =   480
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   8880
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   64
      Top             =   4860
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   9600
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   62
      Top             =   4860
      Width           =   585
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   10320
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   63
      Top             =   4860
      Width           =   585
   End
   Begin VB.PictureBox ITEM3X 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   10320
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   54
      Tag             =   "0"
      Top             =   4035
      Width           =   585
   End
   Begin VB.PictureBox ITEM2X 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   9600
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   58
      Tag             =   "0"
      Top             =   4035
      Width           =   585
   End
   Begin VB.PictureBox ITEM1X 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   8880
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   57
      Tag             =   "0"
      Top             =   4035
      Width           =   585
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   240
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   660
      Width           =   7680
   End
   Begin VB.Image Imageexit 
      Height          =   495
      Left            =   3240
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Imagetrade 
      Height          =   495
      Left            =   2760
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Imagetraining 
      Height          =   495
      Left            =   2040
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Imagespells 
      Height          =   495
      Left            =   1320
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Imageitems 
      Height          =   495
      Left            =   840
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imagestatus 
      Height          =   495
      Left            =   120
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image12 
      Height          =   450
      Left            =   9000
      Top             =   1680
      Width           =   1830
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   9000
      Top             =   2130
      Width           =   1830
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   9000
      Top             =   2580
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   9000
      Top             =   3120
      Width           =   1770
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   79
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9120
      TabIndex        =   78
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9120
      TabIndex        =   77
      Top             =   2250
      Width           =   375
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   76
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label lblSP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000/1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   75
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9120
      TabIndex        =   74
      Top             =   2685
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SSM"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   73
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image9 
      Height          =   615
      Left            =   9600
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   38
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "EVENT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   540
      TabIndex        =   40
      Top             =   3840
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Image HPBAR 
      Height          =   330
      Left            =   9060
      Picture         =   "frmMirage.frx":9368
      Top             =   1770
      Width           =   1725
   End
   Begin VB.Image MPBAR 
      Height          =   330
      Left            =   9060
      Picture         =   "frmMirage.frx":B517
      Top             =   2205
      Width           =   1725
   End
   Begin VB.Image SPBAR 
      Height          =   330
      Left            =   9060
      Picture         =   "frmMirage.frx":D6C6
      Top             =   2640
      Width           =   1725
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySound2 Lib "winmm.dll" Alias "sndPlaySoundB" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Then we have to declare the constants that go along with the sndPlaySound function
Const SND_ASYNC = &H1       'ASYNC allows us to play waves with the ability to interrupt
Const SND_LOOP = &H8        'LOOP causes to sound to be continuously replayed
Const SND_NODEFAULT = &H2   'NODEFAULT causes no sound to be played if the wav can't be found
Const SND_SYNC = &H0        'SYNC plays a wave file without returning control to the calling program until it's finished
Const SND_NOSTOP = &H10     'NOSTOP ensures that we don't stop another wave from playing
Const SND_MEMORY = &H4



Private Sub A_Click()

End Sub

Private Sub Ambiance1_Timer()
'Call AmbS1("\sfx\" & AmbSound1 & ".wav")
End Sub

Private Sub Ambiance2_Timer()
'Call AmbS2("\sfx\" & AmbSound2 & ".wav")
End Sub

Private Sub Ambiance3_Timer()
'Call AmbS3("\sfx\" & AmbSound3 & ".wav")
End Sub

Private Sub Command1_Click()
Call EditorFillLayer
End Sub

Private Sub Command2_Click()
sndPlaySound App.Path & "\sound.wav", SND_SYNC Or SND_NODEFAULT
End Sub

Private Sub Command3_Click()
optGround.value = True
Call cmdClear_Click
optMask.value = True
Call cmdClear_Click
optAnim.value = True
Call cmdClear_Click
optFringe.value = True
Call cmdClear_Click
Call cmdClear2_Click
End Sub

Private Sub Command4_Click()
txtChat.Text = ""
txtActions.Text = ""
txtCombat.Text = ""
frmMirage.SetFocus
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub deathopt_Click()
deathedit.Show vbModal
End Sub

Private Sub fheal_Click()
Healsound.Show vbModal
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://www.silent-shadows.com/ad.php"


 QUICKSPELL1 = 0
 QUICKSPELL2 = 0
 QUICKSPELL3 = 0

'FSOUND_Close

End Sub

Private Sub Form_Resize()
    Call ResizeGUI
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub Option1_Click()
frmQuest.Show vbModal
End Sub

Private Sub Image1_Click()
frmOptions.Show
End Sub

Private Sub Image13_Click()
FrmBlist.Show

End Sub

Private Sub Image16_Click()
picRSM.Visible = False
sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Image17_Click()
'If invbar.value > 0 Then
'If invbar.value < 6 Then
'invbar.value = 0
'Else
'invbar.value = invbar.value - 6
'End If

'Call BlitInventoryX
'End If

End Sub

Private Sub Image18_Click()
'If invbar.value < 41 Then'
'
' invbar.value = invbar.value + 6
' If invbar.value > 41 Then invbar.value = 41

'Call BlitInventoryX
'End If
End Sub

Private Sub Image19_Click()





End Sub

Private Sub Image21_Click()

End Sub

Private Sub Image22_Click()
End Sub

Private Sub Image23_Click()
If Player(MyIndex).WeaponSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).ArmorSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).HelmetSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
    If Player(MyIndex).ShieldSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
SendData ("GIVEITEM" & SEP_CHAR & GetPlayerInvItemNum(MyIndex, frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)) & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
sndPlaySound App.Path & "\SFX\close1.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Image3_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendUseItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value))
End Sub

Private Sub Image4_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Dim value As Long
Dim InvNum As Long

    InvNum = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
    If ITEM1X.Tag = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    ITEM1X.Tag = 0
    ITEM1.Tag = 0
    End If
    If ITEM2X.Tag = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    ITEM2X.Tag = 0
    ITEM2.Tag = 0
    End If
    If ITEM3X.Tag = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    ITEM3X.Tag = 0
    ITEM3.Tag = 0
    End If
    

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value), 0)
        End If
    End If
End Sub

Private Sub Image5_Click()
sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
'picInv.Visible = False
End Sub

Private Sub Image7_Click()
If Player(MyIndex).Spell(frmSpells.lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & frmSpells.lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub Image8_Click()
sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
   'picPlayerSpells.Visible = False
End Sub

Private Sub Image9_Click()
If Player(MyIndex).Access > 0 Then FRMPKLIST.Show

End Sub

Private Sub Inv1_Click()
Call ClearISel
'SI1.Visible = True
'SELectINV = 0

End Sub



Private Sub invbar_Change()

End Sub

Private Sub ISELL_Click()

 If Player(MyIndex).WeaponSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).ArmorSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If
     If Player(MyIndex).HelmetSlot = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value) Then
    Exit Sub
    End If

    
   
    

sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendSellItem(frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value))
End Sub

Private Sub Imageexit_Click()
   sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
    Call GameDestroy
End Sub

Private Sub Imageitems_Click()
    Call UpdateInventory
    frmInventory.LSTINV.ListIndex = 0
    frmInventory.invbar.value = 0
    
    frmInventory.Show
    sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    frmInventory.Show
    
    
End Sub

Private Sub Imagespells_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("spells" & SEP_CHAR & END_CHAR)


End Sub

Private Sub imagestatus_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Imagetrade_Click()
   sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("trade" & SEP_CHAR & END_CHAR)
    
End Sub

Private Sub Imagetraining_Click()
   sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    frmTraining.Show vbModal
    
End Sub

Private Sub ITEM1_Click()
If ITEM1X.Tag = 0 Then Exit Sub
frmMirage.Caption = frmMirage.ITEM1X.Tag
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendUseItem(ITEM1X.Tag)
If Item(ITEM1X.Tag).Type > 3 Then
ITEM1.Tag = 0
ITEM1X.Tag = 0
End If

End Sub

Private Sub ITEM1_DblClick()
ITEM1.Tag = 0
ITEM1X.Tag = 0
End Sub

Private Sub ITEM2_Click()
If ITEM2X.Tag = 0 Then Exit Sub
frmMirage.Caption = frmMirage.ITEM2X.Tag
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendUseItem(ITEM2X.Tag)
If Item(ITEM2X.Tag).Type > 3 Then
ITEM2.Tag = 0
ITEM2X.Tag = 0
End If
End Sub

Private Sub ITEM2_DblClick()
ITEM2.Tag = 0
ITEM2X.Tag = 0
End Sub

Private Sub ITEM3_Click()
If ITEM3X.Tag = 0 Then Exit Sub
frmMirage.Caption = frmMirage.ITEM3X.Tag
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
Call SendUseItem(ITEM3X.Tag)
If Item(ITEM3X.Tag).Type > 3 Then
ITEM3.Tag = 0
ITEM3X.Tag = 0
End If
End Sub

Private Sub ITEM3_DblClick()
ITEM3.Tag = 0
ITEM3X.Tag = 0
End Sub

Private Sub Label12_Click()
Call updateBList
picRSM.Visible = True
    sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    
End Sub

Public Sub quicksl()
tEmp = frmInventory.LSTINV.ListIndex + SELectINV + 1 + (frmInventory.invbar.value)
If (GetPlayerInvItemNum(MyIndex, tEmp) > 0) And (GetPlayerInvItemNum(MyIndex, tEmp) <= MAX_ITEMS) Then
tEmp = GetPlayerInvItemNum(MyIndex, tEmp)
If tEmp = 0 Then Exit Sub
tEmp = frmInventory.LSTINV.ListIndex + SELectINV + (frmInventory.invbar.value)
TEMPITEMNUM = Item(GetPlayerInvItemNum(MyIndex, tEmp + 1)).Pic
TEMPINVSLOT = tEmp
Form2.Label1.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, tEmp + 1)).name)
Form2.Show vbModal
Else

End If
End Sub

Private Sub LSTINV_Click()

End Sub

Private Sub optFishing_Click()
frmFishing.Show vbModal
End Sub

Private Sub optMining_Click()
frmMining.Show vbModal
End Sub

Private Sub optQuest_Click()
frmQuest.Show vbModal
End Sub

Private Sub Optsign_Click()
signedit.Show vbModal
End Sub

Private Sub picInv_Click()

End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call EditorMouseDown(Button, Shift, X, y)
    Call PlayerSearch(Button, Shift, X, y)
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call EditorMouseDown(Button, Shift, X, y)
End Sub

Private Sub SFXopt_Click()
SFXtile.Show vbModal
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckInput(0, KeyCode, Shift)
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If frmMirage.SSTab1.Tab = 0 Then
frmMirage.SSTab1.TabCaption(0) = "chat"
Else
frmMirage.SSTab1.TabCaption(1) = "Global"
End If
frmMirage.picScreen.SetFocus
End Sub

Private Sub SSTab1_GotFocus()
frmMirage.picScreen.SetFocus
End Sub

Private Sub Timer1_Timer()
Call WeaponBltItem
Call ArmorBltItem
Call HelmBltItem
Call SetEquipShizzle
Call ITEMBltItem
Call BlitInventoryX

'Call MusicTest
End Sub



Private Sub txtActions_GotFocus()

frmMirage.picScreen.SetFocus

End Sub

Private Sub txtChat_GotFocus()
    frmMirage.picScreen.SetFocus
End Sub

Private Sub picInventory_Click()
    Call UpdateInventory
    frmInventory.LSTINV.ListIndex = 0
    frmInventory.invbar.value = 0
    
    frmInventory.Show
    sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    frmInventory.Show
    
    
End Sub

Private Sub lblUseItem_Click()
    
End Sub

Private Sub lblDropItem_Click()

End Sub

Private Sub lblCast_Click()
    
End Sub

Private Sub lblCancel_Click()
    
End Sub

Private Sub lblSpellsCancel_Click()
 
End Sub

Private Sub picSpells_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("spells" & SEP_CHAR & END_CHAR)

End Sub

Private Sub picStats_Click()
sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
    
End Sub

Private Sub picTrain_Click()
    sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    frmTraining.Show vbModal
    
End Sub

Private Sub picTrade_Click()
    sndPlaySound App.Path & "\SFX\item1.wav", SND_ASYNC Or SND_NODEFAULT
    Call SendData("trade" & SEP_CHAR & END_CHAR)
    
End Sub

Private Sub picQuit_Click()
    sndPlaySound App.Path & "\SFX\open2.wav", SND_ASYNC Or SND_NODEFAULT
    Call GameDestroy
End Sub

' // MAP EDITOR STUFF //

Private Sub optLayers_Click()
    If optLayers.value = True Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call EditorChooseTile(Button, Shift, X, y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call EditorChooseTile(Button, Shift, X, y)
End Sub

Private Sub cmdSend_Click()
    Call EditorSend
End Sub

Private Sub cmdCancel_Click()
    Call EditorCancel
End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub
Public Function OpenInternet(URL As String)
    Dim ret
    Dim StartURL As String
    StartURL = "start " & URL
    ret = Shell(StartURL, vbHide)
End Function

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub


Private Sub txtCombat_GotFocus()
frmMirage.picScreen.SetFocus
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
' Create new instance of form

 '   Cancel = False
  '  frmBrowser.Show
   'Set ppDisp = frmBrowser.brwWebBrowser.object
   

    'ppDisp = frmBrowser.brwWebBrowser.object

   

End Sub

