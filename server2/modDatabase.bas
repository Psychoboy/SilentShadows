Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Const START_MAP = 1
Public Const START_X = 8
Public Const START_Y = 5

Public Const ADMIN_LOG = "admin.txt"
Public Const PLAYER_LOG = "player.txt"

Function GetVar(File As String, Header As String, Var As String) As String

Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Function FileExist(ByVal FileName As String) As Boolean

    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub SavePlayer(ByVal index As Long)

Dim FileName As String
Dim i As Long
Dim n As Long

    FileName = App.Path & "\accounts\" & Trim(player(index).Login) & ".ini"
    
    Call PutVar(FileName, "GENERAL", "Login", Trim(player(index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim(player(index).Password))
    Call PutVar(FileName, "GENERAL", "Global", Trim(player(index).GlobalPriv))

    For i = 1 To MAX_CHARS
        ' General
        Call PutVar(FileName, "CHAR" & i, "Name", Trim(player(index).Char(i).Name))
        Call PutVar(FileName, "CHAR" & i, "Class", STR(player(index).Char(i).Class))
        Call PutVar(FileName, "CHAR" & i, "Race", STR(player(index).Char(i).Race))
        Call PutVar(FileName, "CHAR" & i, "Sex", STR(player(index).Char(i).Sex))
        Call PutVar(FileName, "CHAR" & i, "Sprite", STR(player(index).Char(i).SPRITE))
        Call PutVar(FileName, "CHAR" & i, "Sprite2", STR(player(index).Char(i).SPRITE2))
        Call PutVar(FileName, "CHAR" & i, "Sprite3", STR(player(index).Char(i).SPRITE3))
        Call PutVar(FileName, "CHAR" & i, "Sprite4", STR(player(index).Char(i).SPRITE4))
        Call PutVar(FileName, "CHAR" & i, "PET", STR(player(index).Char(i).PET))
        
        Call PutVar(FileName, "CHAR" & i, "Level", STR(player(index).Char(i).Level))
        Call PutVar(FileName, "CHAR" & i, "Exp", STR(player(index).Char(i).exp))
        Call PutVar(FileName, "CHAR" & i, "Access", STR(player(index).Char(i).Access))
        Call PutVar(FileName, "CHAR" & i, "PK", STR(player(index).Char(i).PK))
        Call PutVar(FileName, "CHAR" & i, "Guild", Trim(player(index).Char(i).Guild))
        Call PutVar(FileName, "CHAR" & i, "GuildRank", STR(player(index).Char(i).GuildRank))
        Call PutVar(FileName, "CHAR" & i, "Faction", STR(player(index).Char(i).faction))
        ' Vitals
        Call PutVar(FileName, "CHAR" & i, "HP", STR(player(index).Char(i).HP))
        Call PutVar(FileName, "CHAR" & i, "MP", STR(player(index).Char(i).MP))
        Call PutVar(FileName, "CHAR" & i, "SP", STR(player(index).Char(i).SP))
        
        ' Stats
        Call PutVar(FileName, "CHAR" & i, "STR", STR(player(index).Char(i).STR))
        Call PutVar(FileName, "CHAR" & i, "DEF", STR(player(index).Char(i).DEF))
        Call PutVar(FileName, "CHAR" & i, "SPEED", STR(player(index).Char(i).SPeed))
        Call PutVar(FileName, "CHAR" & i, "MAGI", STR(player(index).Char(i).MAGI))
        Call PutVar(FileName, "CHAR" & i, "POINTS", STR(player(index).Char(i).POINTS))
        Call PutVar(FileName, "CHAR" & i, "Crafting", STR(player(index).Char(i).Crafting))
        Call PutVar(FileName, "CHAR" & i, "Fishing", STR(player(index).Char(i).Fishing))
        Call PutVar(FileName, "CHAR" & i, "Mining", STR(player(index).Char(i).Mining))
        
        Call PutVar(FileName, "CHAR" & i, "ANONYMOUS", STR(player(index).Char(i).ANONYMOUS))
        ' Worn equipment
        Call PutVar(FileName, "CHAR" & i, "ArmorSlot", STR(player(index).Char(i).ArmorSlot))
        Call PutVar(FileName, "CHAR" & i, "WeaponSlot", STR(player(index).Char(i).WeaponSlot))
        Call PutVar(FileName, "CHAR" & i, "HelmetSlot", STR(player(index).Char(i).HelmetSlot))
        Call PutVar(FileName, "CHAR" & i, "ShieldSlot", STR(player(index).Char(i).ShieldSlot))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If player(index).Char(i).MAP = 0 Then
            player(index).Char(i).MAP = START_MAP
            player(index).Char(i).X = START_X
            player(index).Char(i).y = START_Y
        End If
        
        If player(index).Char(i).BindMap = 0 Then
            player(index).Char(i).BindMap = START_MAP
            player(index).Char(i).BindX = START_X
            player(index).Char(i).BindY = START_Y
        End If
            
        ' Position
        Call PutVar(FileName, "CHAR" & i, "Map", STR(player(index).Char(i).MAP))
        Call PutVar(FileName, "CHAR" & i, "X", STR(player(index).Char(i).X))
        Call PutVar(FileName, "CHAR" & i, "Y", STR(player(index).Char(i).y))
        Call PutVar(FileName, "CHAR" & i, "Dir", STR(player(index).Char(i).Dir))
        
        Call PutVar(FileName, "CHAR" & i, "BindMap", STR(player(index).Char(i).BindMap))
        Call PutVar(FileName, "CHAR" & i, "BindX", STR(player(index).Char(i).BindX))
        Call PutVar(FileName, "CHAR" & i, "BindY", STR(player(index).Char(i).BindY))
        
        'options
        Call PutVar(FileName, "CHAR" & i, "ooc", STR(player(index).Char(i).ooc))
        Call PutVar(FileName, "CHAR" & i, "tc", STR(player(index).Char(i).tc))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & i, "InvItemNum" & n, STR(player(index).Char(i).Inv(n).Num))
            Call PutVar(FileName, "CHAR" & i, "InvItemVal" & n, STR(player(index).Char(i).Inv(n).Value))
            Call PutVar(FileName, "CHAR" & i, "InvItemDur" & n, STR(player(index).Char(i).Inv(n).Dur))
        Next n
        
        For n = 1 To MAX_BANK
            Call PutVar(FileName, "CHAR" & i, "BankNum" & n, STR(player(index).Char(i).Bank(n).Num))
            Call PutVar(FileName, "CHAR" & i, "BankValue" & n, STR(player(index).Char(i).Bank(n).Value))
            Call PutVar(FileName, "CHAR" & i, "BankDur" & n, STR(player(index).Char(i).Bank(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & i, "Spell" & n, STR(player(index).Char(i).spell(n)))
        Next n
    Next i
End Sub
Sub AddGuild(ByVal GuildName As String, ByVal index As Long)

Dim f As Long
Call SetPlayerGuild(index, GuildName)
Call SetPlayerGuildRank(index, 4)
Call SavePlayer(index)
f = FreeFile
        Open App.Path & "\accounts\guildlist.txt" For Append As #f
            Print #f, GuildName
        Close #f
'Call PlayerMsg(Index, "Guild has been created!", BrightRed)
End Sub
Sub LeaveGuild(ByVal index As Long)

Dim GuildName As String
GuildName = ""
Call SetPlayerGuild(index, GuildName)
Call SetPlayerGuildRank(index, 0)
Call SavePlayer(index)
End Sub
Function FindGuild(ByVal Name As String) As Boolean

Dim f As Long
Dim s As String

    FindGuild = False
    
    f = FreeFile
    Open App.Path & "\accounts\guildlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim(LCase(s)) = Trim(LCase(Name)) Then
                FindGuild = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function
Sub LoadAdminCmds(ByVal index As Long)

Dim FileName As String
FileName = App.Path & "\" & "admincmds.ini"
player(index).Char(player(index).charnum).AdminCmds = GetVar(FileName, STR(player(index).Char(player(index).charnum).Access), "cmds")
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
Dim FileName As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(index)
    
    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"

    player(index).Login = GetVar(FileName, "GENERAL", "Login")
    player(index).Password = GetVar(FileName, "GENERAL", "Password")
    player(index).GlobalPriv = GetVar(FileName, "GENERAL", "Global")
    For i = 1 To MAX_CHARS
        ' General
        player(index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        player(index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        player(index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        player(index).Char(i).Race = Val(GetVar(FileName, "CHAR" & i, "Race"))
        player(index).Char(i).SPRITE = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        player(index).Char(i).SPRITE2 = Val(GetVar(FileName, "CHAR" & i, "Sprite2"))
        player(index).Char(i).SPRITE3 = Val(GetVar(FileName, "CHAR" & i, "Sprite3"))
        player(index).Char(i).SPRITE4 = Val(GetVar(FileName, "CHAR" & i, "Sprite4"))
        player(index).Char(i).PET = Val(GetVar(FileName, "CHAR" & i, "PET"))
        player(index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        player(index).Char(i).exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        player(index).Char(i).Access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        player(index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        player(index).Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
        player(index).Char(i).GuildRank = Val(GetVar(FileName, "CHAR" & i, "GuildRank"))
        player(index).Char(i).faction = Val(GetVar(FileName, "Char" & i, "Faction"))
        If player(index).Char(i).PET = 0 Then player(index).Char(i).PET = 4000
        
        ' Vitals
        player(index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        player(index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        player(index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
        ' Stats
        player(index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        player(index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        player(index).Char(i).SPeed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        player(index).Char(i).MAGI = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        player(index).Char(i).Crafting = Val(GetVar(FileName, "CHAR" & i, "Crafting"))
        player(index).Char(i).Fishing = Val(GetVar(FileName, "CHAR" & i, "Fishing"))
        player(index).Char(i).Mining = Val(GetVar(FileName, "CHAR" & i, "Mining"))
        player(index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        player(index).Char(i).ANONYMOUS = Val(GetVar(FileName, "CHAR" & i, "ANONYMOUS"))
        ' Worn equipment
        player(index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        player(index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        player(index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        player(index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        
        ' Position
        player(index).Char(i).MAP = Val(GetVar(FileName, "CHAR" & i, "Map"))
        player(index).Char(i).X = Val(GetVar(FileName, "CHAR" & i, "X"))
        player(index).Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        player(index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
        ' Bind Position
        player(index).Char(i).BindMap = Val(GetVar(FileName, "CHAR" & i, "BindMap"))
        player(index).Char(i).BindX = Val(GetVar(FileName, "CHAR" & i, "BindX"))
        player(index).Char(i).BindY = Val(GetVar(FileName, "CHAR" & i, "BindY"))
        
        'Options
        ' Options
        player(index).Char(i).ooc = Val(GetVar(FileName, "CHAR" & i, "ooc"))
        ' Options
        player(index).Char(i).tc = Val(GetVar(FileName, "CHAR" & i, "tc"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If player(index).Char(i).MAP = 0 Then
            player(index).Char(i).MAP = START_MAP
            player(index).Char(i).X = START_X
            player(index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            player(index).Char(i).Inv(n).Num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            player(index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            player(index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        For n = 1 To MAX_BANK
            player(index).Char(i).Bank(n).Num = Val(GetVar(FileName, "CHAR" & i, "BankNum" & n))
            player(index).Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankValue" & n))
            player(index).Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            player(index).Char(i).spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n
    Next i
End Sub

Function AccountExist(ByVal Name As String) As Boolean

Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal index As Long, ByVal charnum As Long) As Boolean

    If Trim(player(index).Char(charnum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean

Dim FileName As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"
        RightPassword = GetVar(FileName, "GENERAL", "Password")
        
        If UCase(Trim(Password)) = UCase(Trim(RightPassword)) Then
            PasswordOK = True
        End If
    End If
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)

Dim i As Long

    player(index).Login = Name
    player(index).Password = Password
    player(index).GlobalPriv = 1
    
    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i
    
    Call SavePlayer(index)
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal RaceNum As Byte, ByVal charnum As Long, ByVal HeAd As Integer)

Dim f As Long

    If Trim(player(index).Char(charnum).Name) = "" Then
        player(index).charnum = charnum
        
        player(index).Char(charnum).Name = Name
        player(index).Char(charnum).Sex = Sex
        player(index).Char(charnum).Class = ClassNum
        player(index).Char(charnum).Race = RaceNum
        
        If player(index).Char(charnum).Sex = SEX_MALE Then
            player(index).Char(charnum).SPRITE = 0
        Else
            player(index).Char(charnum).SPRITE = 1
        End If
        
        If player(index).Char(charnum).Sex = SEX_MALE Then
            player(index).Char(charnum).SPRITE2 = HeAd + 533
        Else
            player(index).Char(charnum).SPRITE2 = HeAd + 549
        End If
        
            player(index).Char(charnum).SPRITE3 = 4000
            player(index).Char(charnum).SPRITE4 = 4000

        player(index).Char(charnum).Level = 1
        player(index).Char(charnum).Guild = ""
        player(index).Char(charnum).faction = 0
                    
        player(index).Char(charnum).STR = Class(ClassNum).STR + Race(RaceNum).STR
        player(index).Char(charnum).DEF = Class(ClassNum).DEF + Race(RaceNum).DEF
        player(index).Char(charnum).SPeed = Class(ClassNum).SPeed + Race(RaceNum).SPeed
        player(index).Char(charnum).MAGI = Class(ClassNum).MAGI + Race(RaceNum).MAGI
        
        player(index).Char(charnum).MAP = START_MAP
        player(index).Char(charnum).X = START_X
        player(index).Char(charnum).y = START_Y
        
        player(index).Char(charnum).BindMap = START_MAP
        player(index).Char(charnum).BindX = START_X
        player(index).Char(charnum).BindY = START_Y
        player(index).Char(charnum).ooc = 1
        player(index).Char(charnum).tc = 1
            
        player(index).Char(charnum).HP = GetPlayerMaxHP(index)
        player(index).Char(charnum).MP = GetPlayerMaxMP(index)
        player(index).Char(charnum).SP = GetPlayerMaxSP(index)
                
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        
        Call SavePlayer(index)
            
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal index As Long, ByVal charnum As Long)

Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(player(index).Char(charnum).Name)
    Call ClearChar(index, charnum)
    Call SavePlayer(index)
End Sub

Function FindChar(ByVal Name As String) As Boolean

Dim f As Long
Dim s As String

    FindChar = False
    
    f = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim(LCase(s)) = Trim(LCase(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
On Error Resume Next

Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
            DoEvents
        End If
    Next i
End Sub

Sub LoadVersion()
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\version.ini"
    
    ' Check if file exists
    If Not FileExist("version.ini") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
  
        CLIENT_MAJOR = Val(Trim(GetVar(FileName, "VERSION", "Major")))
        CLIENT_MINOR = Val(Trim(GetVar(FileName, "VERSION", "Minor")))
        CLIENT_REVISION = Val(Trim(GetVar(FileName, "VERSION", "Revision")))
End Sub

Sub LoadClasses()
Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).SPRITE = GetVar(FileName, "CLASS" & i, "Sprite")
        Class(i).SPRITE2 = GetVar(FileName, "CLASS" & i, "Sprite2")
        Class(i).SPRITE3 = GetVar(FileName, "CLASS" & i, "Sprite3")
        Class(i).SPRITE4 = GetVar(FileName, "CLASS" & i, "Sprite4")
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPeed = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses()

Dim FileName As String
Dim i As Long

    FileName = App.Path & "\classes.ini"
    
    For i = 0 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(i).SPRITE))
        Call PutVar(FileName, "CLASS" & i, "Sprite2", STR(Class(i).SPRITE2))
        Call PutVar(FileName, "CLASS" & i, "Sprite3", STR(Class(i).SPRITE3))
        Call PutVar(FileName, "CLASS" & i, "Sprite4", STR(Class(i).SPRITE4))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPeed))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub LoadRaces()
Dim FileName As String
Dim i As Long

    Call CheckRaces
    
    FileName = App.Path & "\races.ini"
    
    Max_Races = Val(GetVar(FileName, "INIT", "MaxRaces"))
    
    ReDim Race(0 To Max_Races) As RaceRec
    
    Call ClearRaces
    
    For i = 0 To Max_Races
        Race(i).Name = GetVar(FileName, "RACE" & i, "Name")
        Race(i).SPRITE = GetVar(FileName, "RACE" & i, "Sprite")
        Race(i).SPRITE2 = GetVar(FileName, "RACE" & i, "Sprite2")
        Race(i).SPRITE3 = GetVar(FileName, "RACE" & i, "Sprite3")
        Race(i).SPRITE4 = GetVar(FileName, "RACE" & i, "Sprite4")
        Race(i).STR = Val(GetVar(FileName, "RACE" & i, "STR"))
        Race(i).DEF = Val(GetVar(FileName, "RACE" & i, "DEF"))
        Race(i).SPeed = Val(GetVar(FileName, "RACE" & i, "SPEED"))
        Race(i).MAGI = Val(GetVar(FileName, "RACE" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub SaveRaces()

Dim FileName As String
Dim i As Long

    FileName = App.Path & "\races.ini"
    
    For i = 0 To Max_Races
        Call PutVar(FileName, "RACE" & i, "Name", Trim(Race(i).Name))
        Call PutVar(FileName, "RACE" & i, "Sprite", STR(Race(i).SPRITE))
        Call PutVar(FileName, "RACE" & i, "Sprite2", STR(Race(i).SPRITE2))
        Call PutVar(FileName, "RACE" & i, "Sprite3", STR(Race(i).SPRITE3))
        Call PutVar(FileName, "RACE" & i, "Sprite4", STR(Race(i).SPRITE4))
        Call PutVar(FileName, "RACE" & i, "STR", STR(Race(i).STR))
        Call PutVar(FileName, "RACE" & i, "DEF", STR(Race(i).DEF))
        Call PutVar(FileName, "RACE" & i, "SPEED", STR(Race(i).SPeed))
        Call PutVar(FileName, "RACE" & i, "MAGI", STR(Race(i).MAGI))
    Next i
End Sub


Sub CheckRaces()
    If Not FileExist("races.ini") Then
        Call SaveRaces
    End If
End Sub

Sub SaveItems()

Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)

Dim FileName As String

    FileName = App.Path & "\items.ini"
    
    Call PutVar(FileName, "ITEM" & ItemNum, "Name", Trim(Item(ItemNum).Name))
    Call PutVar(FileName, "ITEM" & ItemNum, "Pic", Trim(Item(ItemNum).Pic))
    Call PutVar(FileName, "ITEM" & ItemNum, "Type", Trim(Item(ItemNum).Type))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data1", Trim(Item(ItemNum).Data1))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data2", Trim(Item(ItemNum).Data2))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data3", Trim(Item(ItemNum).Data3))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data4", Trim(Item(ItemNum).Data4))
    Call PutVar(FileName, "ITEM" & ItemNum, "CLass", Trim(Item(ItemNum).Class))
    Call PutVar(FileName, "ITEM" & ItemNum, "STRmod", Trim(Item(ItemNum).STRmod))
    Call PutVar(FileName, "ITEM" & ItemNum, "DEFmod", Trim(Item(ItemNum).DEFmod))
    Call PutVar(FileName, "ITEM" & ItemNum, "MAGImod", Trim(Item(ItemNum).MAGImod))
    Call PutVar(FileName, "ITEM" & ItemNum, "SPRITE", Trim(Item(ItemNum).SPRITE))
    Call PutVar(FileName, "ITEM" & ItemNum, "SellValue", Trim(Item(ItemNum).SellValue))
    Call PutVar(FileName, "ITEM" & ItemNum, "NoDrop", Trim(Item(ItemNum).NoDrop))
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long

    Call CheckItems
    
    FileName = App.Path & "\items.ini"
    
    For i = 1 To MAX_ITEMS
        Item(i).Name = GetVar(FileName, "ITEM" & i, "Name")
        Item(i).Pic = Val(GetVar(FileName, "ITEM" & i, "Pic"))
        Item(i).Type = Val(GetVar(FileName, "ITEM" & i, "Type"))
        Item(i).Data1 = Val(GetVar(FileName, "ITEM" & i, "Data1"))
        Item(i).Data2 = (GetVar(FileName, "ITEM" & i, "Data2"))
        Item(i).Data3 = Val(GetVar(FileName, "ITEM" & i, "Data3"))
        Item(i).Data4 = Val(GetVar(FileName, "ITEM" & i, "Data4"))
        Item(i).Class = Val(GetVar(FileName, "ITEM" & i, "CLass"))
        Item(i).STRmod = Val(GetVar(FileName, "ITEM" & i, "STRmod"))
        Item(i).DEFmod = Val(GetVar(FileName, "ITEM" & i, "DEFmod"))
        Item(i).MAGImod = Val(GetVar(FileName, "ITEM" & i, "MAGImod"))
        Item(i).SPRITE = Val(GetVar(FileName, "ITEM" & i, "SPRITE"))
        Item(i).SellValue = Val(GetVar(FileName, "ITEM" & i, "SellValue"))
        'Item(i).NoDrop = 0
        Item(i).NoDrop = Val(GetVar(FileName, "ITEM" & i, "NoDrop"))
        
        
        DoEvents
    Next i
End Sub

Sub CheckItems()
    If Not FileExist("items.ini") Then
        Call SaveItems
    End If
End Sub

Sub SaveShops()

Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)

Dim FileName As String
Dim i As Long

    FileName = App.Path & "\shops.ini"
    
    Call PutVar(FileName, "SHOP" & ShopNum, "Name", Trim(Shop(ShopNum).Name))
    Call PutVar(FileName, "SHOP" & ShopNum, "JoinSay", Trim(Shop(ShopNum).JoinSay))
    Call PutVar(FileName, "SHOP" & ShopNum, "LeaveSay", Trim(Shop(ShopNum).LeaveSay))
    Call PutVar(FileName, "SHOP" & ShopNum, "FixesItems", Trim(Shop(ShopNum).FixesItems))
    Call PutVar(FileName, "SHOP" & ShopNum, "OneSale", Trim(Shop(ShopNum).OneSale))
    
    For i = 1 To MAX_TRADES
        Call PutVar(FileName, "SHOP" & ShopNum, "GiveItem" & i, Trim(Shop(ShopNum).TradeItem(i).GiveItem))
        Call PutVar(FileName, "SHOP" & ShopNum, "GiveValue" & i, Trim(Shop(ShopNum).TradeItem(i).GiveValue))
        Call PutVar(FileName, "SHOP" & ShopNum, "GetItem" & i, Trim(Shop(ShopNum).TradeItem(i).GetItem))
        Call PutVar(FileName, "SHOP" & ShopNum, "GetValue" & i, Trim(Shop(ShopNum).TradeItem(i).GetValue))
    Next i
End Sub

Sub LoadShops()
On Error Resume Next
'leave
Dim FileName As String
Dim X As Long, y As Long

    Call CheckShops
    
    FileName = App.Path & "\shops.ini"
    
    For y = 1 To MAX_SHOPS
        Shop(y).Name = GetVar(FileName, "SHOP" & y, "Name")
        Shop(y).JoinSay = GetVar(FileName, "SHOP" & y, "JoinSay")
        Shop(y).LeaveSay = GetVar(FileName, "SHOP" & y, "LeaveSay")
        Shop(y).FixesItems = GetVar(FileName, "SHOP" & y, "FixesItems")
        Shop(y).OneSale = GetVar(FileName, "SHOP" & y, "OneSale")
        
        For X = 1 To MAX_TRADES
            Shop(y).TradeItem(X).GiveItem = GetVar(FileName, "SHOP" & y, "GiveItem" & X)
            Shop(y).TradeItem(X).GiveValue = GetVar(FileName, "SHOP" & y, "GiveValue" & X)
            Shop(y).TradeItem(X).GetItem = GetVar(FileName, "SHOP" & y, "GetItem" & X)
            Shop(y).TradeItem(X).GetValue = GetVar(FileName, "SHOP" & y, "GetValue" & X)
        Next X
    
        DoEvents
    Next y
End Sub

Sub CheckShops()
    If Not FileExist("shops.ini") Then
        Call SaveShops
    End If
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\spells.ini"
    
    Call PutVar(FileName, "SPELL" & SpellNum, "Name", Trim(spell(SpellNum).Name))
    Call PutVar(FileName, "SPELL" & SpellNum, "LvlReq", Trim(spell(SpellNum).LevelReq))
    Call PutVar(FileName, "SPELL" & SpellNum, "ClassReq", Trim(spell(SpellNum).ClassReq))
    Call PutVar(FileName, "SPELL" & SpellNum, "Type", Trim(spell(SpellNum).Type))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data1", Trim(spell(SpellNum).Data1))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data2", Trim(spell(SpellNum).Data2))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data3", Trim(spell(SpellNum).Data3))
    Call PutVar(FileName, "SPELL" & SpellNum, "MPused", Trim(spell(SpellNum).MPused))
    Call PutVar(FileName, "SPELL" & SpellNum, "Gfx", Trim(spell(SpellNum).Gfx))
    Call PutVar(FileName, "SPELL" & SpellNum, "Sfx", Trim(spell(SpellNum).Sfx))

End Sub

Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next i
End Sub

Sub LoadSpells()
Dim FileName As String
Dim i As Long

    Call CheckSpells
    
    FileName = App.Path & "\spells.ini"
    
    For i = 1 To MAX_SPELLS
        spell(i).Name = GetVar(FileName, "SPELL" & i, "Name")
        spell(i).ClassReq = Val(GetVar(FileName, "SPELL" & i, "ClassReq"))
        spell(i).LevelReq = Val(GetVar(FileName, "SPELL" & i, "LvlReq"))
        spell(i).Type = Val(GetVar(FileName, "SPELL" & i, "Type"))
        spell(i).Data1 = Val(GetVar(FileName, "SPELL" & i, "Data1"))
        spell(i).Data2 = Val(GetVar(FileName, "SPELL" & i, "Data2"))
        spell(i).Data3 = Val(GetVar(FileName, "SPELL" & i, "Data3"))
        spell(i).MPused = Val(GetVar(FileName, "SPELL" & i, "MPused"))
        spell(i).Gfx = Val(GetVar(FileName, "SPELL" & i, "Gfx"))
        spell(i).Sfx = Val(GetVar(FileName, "SPELL" & i, "Sfx"))
        DoEvents
    Next i
End Sub

Sub CheckSpells()
    If Not FileExist("spells.ini") Then
        Call SaveSpells
    End If
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String

    FileName = App.Path & "\npcs.ini"
    
    Call PutVar(FileName, "NPC" & NpcNum, "Name", Trim(Npc(NpcNum).Name))
    Call PutVar(FileName, "NPC" & NpcNum, "AttackSay", Trim(Npc(NpcNum).AttackSay))
    Call PutVar(FileName, "NPC" & NpcNum, "Sprite", Trim(Npc(NpcNum).SPRITE))
     Call PutVar(FileName, "NPC" & NpcNum, "Sprite2", Trim(Npc(NpcNum).SPRITE2))
      Call PutVar(FileName, "NPC" & NpcNum, "Sprite3", Trim(Npc(NpcNum).SPRITE3))
       Call PutVar(FileName, "NPC" & NpcNum, "Sprite4", Trim(Npc(NpcNum).SPRITE4))
    Call PutVar(FileName, "NPC" & NpcNum, "SpawnSecs", Trim(Npc(NpcNum).SpawnSecs))
    Call PutVar(FileName, "NPC" & NpcNum, "Behavior", Trim(Npc(NpcNum).Behavior))
    Call PutVar(FileName, "NPC" & NpcNum, "Range", Trim(Npc(NpcNum).Range))
    Call PutVar(FileName, "NPC" & NpcNum, "DropChance", Trim(Npc(NpcNum).DropChance))
    Call PutVar(FileName, "NPC" & NpcNum, "DropItem", Trim(Npc(NpcNum).DropItem))
    Call PutVar(FileName, "NPC" & NpcNum, "DropItemValue", Trim(Npc(NpcNum).DropItemValue))
    Call PutVar(FileName, "NPC" & NpcNum, "STR", Trim(Npc(NpcNum).STR))
    Call PutVar(FileName, "NPC" & NpcNum, "DEF", Trim(Npc(NpcNum).DEF))
    Call PutVar(FileName, "NPC" & NpcNum, "SPEED", Trim(Npc(NpcNum).SPeed))
    Call PutVar(FileName, "NPC" & NpcNum, "MAGI", Trim(Npc(NpcNum).MAGI))
    Call PutVar(FileName, "NPC" & NpcNum, "DropChance2", Trim(Npc(NpcNum).DropChance2))
    Call PutVar(FileName, "NPC" & NpcNum, "DropItem2", Trim(Npc(NpcNum).DropItem2))
    Call PutVar(FileName, "NPC" & NpcNum, "DropItemValue2", Trim(Npc(NpcNum).DropItemValue2))
End Sub

Sub LoadNpcs()
On Error Resume Next

Dim FileName As String
Dim i As Long

    Call CheckNpcs
    
    FileName = App.Path & "\npcs.ini"
    
    For i = 1 To MAX_NPCS
        Npc(i).Name = GetVar(FileName, "NPC" & i, "Name")
        Npc(i).AttackSay = GetVar(FileName, "NPC" & i, "AttackSay")
        Npc(i).SPRITE = GetVar(FileName, "NPC" & i, "Sprite")
        Npc(i).SPRITE2 = GetVar(FileName, "NPC" & i, "Sprite2")
        Npc(i).SPRITE3 = GetVar(FileName, "NPC" & i, "Sprite3")
        Npc(i).SPRITE4 = GetVar(FileName, "NPC" & i, "Sprite4")
        Npc(i).SpawnSecs = GetVar(FileName, "NPC" & i, "SpawnSecs")
        Npc(i).Behavior = GetVar(FileName, "NPC" & i, "Behavior")
        Npc(i).Range = GetVar(FileName, "NPC" & i, "Range")
        Npc(i).DropChance = GetVar(FileName, "NPC" & i, "DropChance")
        Npc(i).DropItem = GetVar(FileName, "NPC" & i, "DropItem")
        Npc(i).DropItemValue = GetVar(FileName, "NPC" & i, "DropItemValue")
        Npc(i).STR = GetVar(FileName, "NPC" & i, "STR")
        Npc(i).DEF = GetVar(FileName, "NPC" & i, "DEF")
        Npc(i).SPeed = GetVar(FileName, "NPC" & i, "SPEED")
        Npc(i).MAGI = GetVar(FileName, "NPC" & i, "MAGI")
        Npc(i).DropChance2 = GetVar(FileName, "NPC" & i, "DropChance2")
        Npc(i).DropItem2 = GetVar(FileName, "NPC" & i, "DropItem2")
        Npc(i).DropItemValue2 = GetVar(FileName, "NPC" & i, "DropItemValue2")
    
        DoEvents
    Next i
End Sub

Sub CheckNpcs()
    If Not FileExist("npcs.ini") Then
        Call SaveNpcs
    End If
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , MAP(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next i
End Sub

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , MAP(i)
        Close #f
    
        DoEvents
    Next i
End Sub

Sub ConvertOldMapsToNew()
Dim FileName As String
Dim i As Long
Dim f As Long
Dim X As Long, y As Long
Dim OldMap As OldMapRec
Dim NewMap As MapRec

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & i & ".dat"
        
        ' Get the old file
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , OldMap
        Close #f
        
        ' Delete the old file
        Call Kill(FileName)
        
        ' Convert
        NewMap.Name = OldMap.Name
        NewMap.Revision = OldMap.Revision + 3
        NewMap.Moral = OldMap.Moral
        NewMap.Up = OldMap.Up
        NewMap.Down = OldMap.Down
        NewMap.Left = OldMap.Left
        NewMap.Right = OldMap.Right
        NewMap.Music = OldMap.Music
        NewMap.BootMap = OldMap.BootMap
        NewMap.BootX = OldMap.BootX
        NewMap.BootY = OldMap.BootY
        NewMap.Shop = OldMap.Shop
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                NewMap.Tile(X, y).Ground = OldMap.Tile(X, y).Ground
                NewMap.Tile(X, y).Mask = OldMap.Tile(X, y).Mask
                NewMap.Tile(X, y).Anim = OldMap.Tile(X, y).Anim
                NewMap.Tile(X, y).Mask2 = OldMap.Tile(X, y).Mask2
                NewMap.Tile(X, y).M2Anim = OldMap.Tile(X, y).M2Anim
                NewMap.Tile(X, y).Fringe = OldMap.Tile(X, y).Fringe
                NewMap.Tile(X, y).FAnim = OldMap.Tile(X, y).FAnim
                NewMap.Tile(X, y).Fringe2 = OldMap.Tile(X, y).Fringe2
                NewMap.Tile(X, y).FAnim = OldMap.Tile(X, y).FAnim
                NewMap.Tile(X, y).Type = OldMap.Tile(X, y).Type
                NewMap.Tile(X, y).Data1 = OldMap.Tile(X, y).Data1
                NewMap.Tile(X, y).Data2 = OldMap.Tile(X, y).Data2
                NewMap.Tile(X, y).Data3 = OldMap.Tile(X, y).Data3
                
            Next X
        Next y
        
        For X = 1 To MAX_MAP_NPCS
            NewMap.Npc(X) = OldMap.Npc(X)
        Next X
        ' Set new values to 0 or null
        NewMap.Indoors = NO
        
        
        ' Save the new map
        f = FreeFile
        Open FileName For Binary As #f
            Put #f, , NewMap
        Close #f
    Next i
End Sub

Sub CheckMaps()
Dim FileName As String
Dim X As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
Dim FileName As String
Dim f As Long
FN = Format(Now, "mm-dd-yyyy") & FN
    If ServerLog = True Then
        FileName = App.Path & "\" & FN
    
        If Not FileExist(FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid(IP, i, 1) = "." Then
            Exit For
        End If
    Next i
    IP = Mid(IP, 1, i)
            
    f = FreeFile
    Open FileName For Append As #f
        Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    
    Call GlobalMsgCombat(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub



Sub DeleteName(ByVal Name As String)
Dim f1 As Long, f2 As Long
Dim s As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim(LCase(s)) <> Trim(LCase(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub
