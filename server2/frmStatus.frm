VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   0  'None
   Caption         =   "frmStatus"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckWebStat 
      Index           =   0
      Left            =   150
      Top             =   160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblUT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim charBuffer
Dim charIndex As Integer

Dim iPlayer As Long, iGM As Long

Dim ProgUptime As String
Dim HourDiff As String, MinDiff As String, SecDiff As String

Private fso As New FileSystemObject

Private Sub Form_Load()
Dim i As Integer

With sckWebStat(0)
    .Close
    .LocalPort = GAME_PORT + 1
    .Listen
End With

For i = 1 To 255
    Load sckWebStat(i)
Next i
End Sub

Private Sub sckWebStat_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim i As Integer

If Index = 0 Then
    For i = 1 To 255
        If sckWebStat(i).State = sckClosed Then
            Exit For
         End If
    Next i
      sckWebStat(i).LocalPort = 0
      sckWebStat(i).Accept requestID
End If
End Sub

Private Sub sckWebStat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next

Dim SPC2 As Integer
Dim strData As String, htmlData As String
Dim RequestedPage As String
Dim FindGet As String, PageToGet As String

sckWebStat(Index).GetData strData$

If Mid(strData$, 1, 3) = "GET" Then
    FindGet = InStr(strData, "GET ")
    SPC2 = InStr(FindGet + 5, strData, " ")
    PageToGet = Mid(strData, FindGet + 4, SPC2 - (FindGet + 4))
    RequestedPage = PageToGet
End If


If RequestedPage$ = "/" Then
    RequestedPage$ = "\serverstatus.html"
Else
    RequestedPage$ = "\serverstatus.html" 'Mid(RequestedPage$, 2, Len(RequestedPage$) - 1)
End If

If fso.FileExists(App.Path & "\html" & RequestedPage$) Then
Call CountChars
Call PlayCharStatus
Call GetUptime

    htmlData$ = Read_HTML(App.Path & "\html\serverstatus.html")
    htmlData$ = Replace(htmlData$, "$SERVNAME", GAME_NAME)
    htmlData$ = Replace(htmlData$, "$VERSION", CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION)

    htmlData$ = Replace(htmlData$, "$ACCTCOUNT", CountAccounts(App.Path & "\accounts\*.ini"))
    htmlData$ = Replace(htmlData$, "$CHARCOUNT", charIndex)
    htmlData$ = Replace(htmlData$, "$ITMCOUNT", "<not working>")
    htmlData$ = Replace(htmlData$, "$NPCCOUNT", "<not working>")

If GameTime = 0 Then
    htmlData$ = Replace(htmlData$, "$GAMETIME", "Day")
Else
    htmlData$ = Replace(htmlData$, "$GAMETIME", "Night")
End If

    htmlData$ = Replace(htmlData$, "$ONLINEPLRS", " " & TotalOnlinePlayers & "/" & MAX_PLAYERS & " ")
    htmlData$ = Replace(htmlData$, "$PLAYERONLINE", iPlayer)
    htmlData$ = Replace(htmlData$, "$GMONLINE", iGM)
    
    htmlData$ = Replace(htmlData$, "$UPTIME", ProgUptime)


    sckWebStat(Index).SendData htmlData$ & vbCrLf
End If

End Sub

Private Sub sckWebStat_SendComplete(Index As Integer)
    sckWebStat(Index).Close
End Sub

Public Sub GetUptime()
Dim T1 As String, T2 As String
Dim Diff As String

    T1 = Time()
    T2 = lblUT.Caption
    Diff = TimeValue(T2) - TimeValue(T1)
    HourDiff = Hour(Diff)
    MinDiff = Minute(Diff)
    SecDiff = Second(Diff)
    
    ProgUptime = " " + HourDiff + " Hours, " + MinDiff + " Minutes, " + SecDiff + " Seconds."
End Sub

Public Sub CountChars()
charIndex = 0
Dim f As Long

f = FreeFile
Open App.Path & "\accounts\charlist.txt" For Input As #f
    While Not EOF(f)
        Input #f, charBuffer

        DoEvents
            charIndex = charIndex + 1
    Wend
Close #f
End Sub

Public Sub PlayCharStatus()
Dim i As Long

iPlayer = 0
iGM = 0

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
        If GetPlayerAccess(i) <= 0 Then
            iPlayer = iPlayer + 1
        Else
            iGM = iGM + 1
        End If
    End If
Next i

End Sub

Public Function CountAccounts(Optional sDirectory As String = "*.*") As Long
Dim lCount As Long
'Dim lAttributes As Long

'lAttributes = VBA.vbNormal

If Len(Dir(sDirectory)) <> 0 Then ', lAttributes)) <> 0 Then
    lCount = 1

    Do While Len(Dir) > 0
        lCount = lCount + 1
    Loop
End If

CountAccounts = lCount
    
End Function

Private Function Read_HTML(FileName)
On Error Resume Next

Dim f As Long
Dim HTML As String

HTML = ""

f = FreeFile
If fso.FileExists(FileName) Then
    If Len(FileName) Then
        Open FileName For Binary As #f
            HTML = Input(LOF(f), #f)
           ' DoEvents
        Close #f
    End If
Read_HTML = HTML
Else
Read_HTML = ""
End If

End Function

