Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = "Mirage Server <IP " & frmServer.Socket(0).LocalIP & " Port " & STR(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) And player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    If IsConnected(index) And Trim(player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(player(i).Login)) = LCase(Trim(Login)) Then
            IsMultiAccounts = True
            Call CloseSocket(i)
            Exit Function
        End If
    Next i
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim(GetPlayerIP(i)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Is banned?
        If Trim(LCase(fIP)) = Trim(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub
Sub GlobalMsgCombat(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSGCOMBAT" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub
Sub PlayerMsgCombat(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSGCOMBAT" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub


Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub
Sub MapMsgCombat(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim text As String

    Packet = "PLAYERMSGCOMBAT" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsgCombat(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    If index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(index)) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(index) & ".", True)
            Else
                Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
On Error Resume Next

Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Integer

    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
        
        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Call CloseSocket(index)
        End If
            
        player(index).Buffer = player(index).Buffer & Buffer
        
        Start = InStr(player(index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(player(index).Buffer, 1, Start - 1)
            player(index).Buffer = Mid(player(index).Buffer, Start + 1, Len(player(index).Buffer))
            player(index).DataPackets = player(index).DataPackets + 1
            Start = InStr(player(index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If
        Loop
                
        ' Check if elapsed time has passed
        player(index).DataBytes = player(index).DataBytes + DataLength
        If GetTickCount >= player(index).DataTimer + 1000 Then
            player(index).DataTimer = GetTickCount
            player(index).DataBytes = 0
            player(index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If player(index).DataBytes > 1000 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If player(index).DataPackets > 25 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
End Sub

Sub HandleData(ByVal index As Long, ByVal Data As String)
On Error Resume Next

Dim parse() As String
Dim Name As String
Dim Temp As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim Race As Long
Dim charnum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim i As Long, n As Long, X As Long, y As Long, f As Long
Dim percentage As String * 4
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim GuildName As String
Dim GuildInvite As String
Dim NewRank As Byte
Dim nPlayer As Long

        
    ' Handle Data
    parse = Split(Data, SEP_CHAR)
        
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "getclasses" Then
        If Not IsPlaying(index) Then
            Call SendNewCharClasses(index)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting races for making a character   ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "getraces" Then
        If Not IsPlaying(index) Then
            Call SendNewCharRaces(index)
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "newaccount" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
               
            ' Get the data
            Name = parse(1)
            Password = parse(2)
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If InStr(Name, " ") Then
                Call AlertMsg(index, "Your account name can not contain spaces")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.", True)
                'Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(index, "Your account has been created!")
                Call AlertMsg(index, "Your account has been created!")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Delete account packet ::
    ' :::::::::::::::::::::::::::
    If LCase(parse(0)) = "delaccount" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
            ' Get the data
            Name = parse(1)
            Password = parse(2)
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(index, Name)
            For i = 1 To MAX_CHARS
                If Trim(player(index).Char(i).Name) <> "" Then
                    Call DeleteName(player(index).Char(i).Name)
                End If
            Next i
            Call ClearPlayer(index)
            
            ' Everything went ok
            Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
            'Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    If LCase(parse(0)) = "login" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
            ' Get the data
            Name = parse(1)
            Password = parse(2)
            
            If InStr(Name, "/") Then
                Call AlertMsg(index, "Your account name can not contain /")
                Exit Sub
            End If
            
            If InStr(Name, "\") Then
                Call AlertMsg(index, "Your account name can not contain \")
                Exit Sub
            End If
        
        Name = XORDecrypt(Name, "dfg8759j35")
    Password = XORDecrypt(Password, "dfg87dgkjh345")
        
            ' Check versions
             
             If Val(parse(3)) < CLIENT_MAJOR Or Val(parse(4)) < CLIENT_MINOR Or Val(parse(5)) < CLIENT_REVISION Then
                Call AlertMsg(index, "Version is Outdated.")
                Exit Sub
            End If
            
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
For i = 1 To Len(Name)
    n = Asc(Mid(Name, i, 1))

    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
    Else
        Call AlertMsg(index, "Invalid name, only letters, numbers, and _ allowed in names.")
        Exit Sub
    End If
Next i
            
            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If
        
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If
        
            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized. Closing other Connection. Try Again.")
                
                Exit Sub
            End If
                
            ' Everything went ok
    
            ' Load the player
            Call LoadPlayer(index, Name)
            'Call LoadPlayer(Index, Name)
            Call SendChars(index)
            Call SendChars(index)
    
            ' Show the player up on the socket status
            'Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
        End If
        Exit Sub
    End If
    
    If LCase(parse(0)) = "anonymous" Then
     Call SetPlayeranonymous(index, "1")
     Exit Sub
    End If
    
    If LCase(parse(0)) = "joinfaction" Then
     If GetPlayerFaction(index) > 0 Then
      Call PlayerMsg(index, "Your already in a faction", BrightRed)
      Exit Sub
     End If
     If GetPlayerLevel(index) < 15 Then
      Call PlayerMsg(index, "You must be at least level 15 before joining a faction!", BrightRed)
      Exit Sub
     End If
     If MAP(GetPlayerMap(index)).BootMap <= 1 Or MAP(GetPlayerMap(index)).BootMap >= 4 Then
      Call PlayerMsg(index, "you cannot join a faction here", BrightRed)
      Exit Sub
     End If
     n = MAP(GetPlayerMap(index)).BootMap
     Call SetPlayerFaction(index, n)
     Call PlayerMsg(index, "You have joined the faction!", BrightRed)
     Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Start Guild Packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "startguild" Then
    GuildName = parse(1)
    If GetPlayerGuildRank(index) > 0 Then
    Call PlayerMsgCombat(index, "You are already in a guild.", BrightRed)
    Exit Sub
    End If
    If GetPlayerLevel(index) < 20 Then
    Call PlayerMsgCombat(index, "Your level is to low to make a guild. You must be at least level 20.", BrightRed)
    Exit Sub
    End If
    
    For i = 1 To Len(GuildName)
              n = Asc(Mid(GuildName, i, 1))
              
              If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
              Else
                  Call PlayerMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", BrightRed)
                  Exit Sub
              End If
    Next i
    If FindGuild(GuildName) Then
                Call PlayerMsg(index, "Sorry, but that Guild Name is in use!", BrightRed)
                Exit Sub
            End If
            
    If HasItem(index, 984) >= 1 Then
            Call TakeItem(index, 984, 1)
    Else
        Call PlayerMsgCombat(index, "Starting a Guild requires a guild ticket", BrightRed)
        Exit Sub
    End If
    
    Call AddGuild(GuildName, index)
    Call PlayerMsgCombat(index, "Guild has been created!", BrightRed)
    Call SavePlayer(index)
    Exit Sub
    End If
    ' :::::::::::::::::
    ' :: Leave Guild ::
    ' :::::::::::::::::
    If LCase(parse(0)) = "leaveguild" Then
    Call LeaveGuild(index)
    Call PlayerMsgCombat(index, "You have left the guild!", BrightRed)
    Call SavePlayer(index)
    Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Invite Guild ::
    ' ::::::::::::::::::
    If LCase(parse(0)) = "guildinvite" Then
    n = FindPlayer(parse(1))
        If n = index Then
            Exit Sub
        End If
        If GetPlayerGuildRank(index) <= 1 Then
        Call PlayerMsgCombat(index, "You do not have permission to invite any player to the guild!", Blue)
        Exit Sub
        End If
        If GetPlayerGuildRank(n) > 0 Then
            Call PlayerMsgCombat(index, "The player is already in a guild", BrightRed)
            Exit Sub
        End If
        
        Call PlayerMsgCombat(n, "You have been invited to join " & GetPlayerGuild(index) & " Type /guildjoin to join the guild", Blue)
        Call PlayerMsgCombat(index, "You have sent the invitation.", Blue)
        player(n).InviteGuild = index
        Exit Sub
    End If
    
    ' ::::::::::::::::
    ' :: Join Guild ::
    ' ::::::::::::::::
    If LCase(parse(0)) = "guildjoin" Then
    n = player(index).InviteGuild
    If GetPlayerGuildRank(index) > 0 Then
        Call PlayerMsgCombat(index, "You are already in a guild", Blue)
        Exit Sub
    End If
    Call SetPlayerGuild(index, GetPlayerGuild(n))
    Call SetPlayerGuildRank(index, 1)
    Call PlayerMsgCombat(index, "You have joined a guild!", Blue)
    s = "<guild>: " & GetPlayerName(index) & " has joined the guild!"
    'Call AddLog(s, PLAYER_LOG)
    Call GlobalMsgGuild(index, s, Blue)
    Call SavePlayer(n)
    Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Set Guild Rank ::
    ' ::::::::::::::::::::
    
    If LCase(parse(0)) = "setguildrank" Then
    n = FindPlayer(parse(1))
    NewRank = Val(parse(2))
    nPlayer = parse(1)
    
    If GetPlayerGuildRank(index) > GetPlayerGuildRank(n) And GetPlayerGuildRank(index) > 2 Then
        If GetPlayerGuild(n) = GetPlayerGuild(index) And GetPlayerGuildRank(n) > 0 Then
            If NewRank > 0 Then
            Call SetPlayerGuildRank(n, NewRank)
            Call PlayerMsgCombat(n, "Your Guild Rank is now " & NewRank, Yellow)
            Call PlayerMsgCombat(index, "You have set " & GetPlayerName(n) & " Rank to " & NewRank & ".", BrightBlue)
            Call SavePlayer(n)
            Else
            Call SetPlayerGuildRank(n, NewRank)
            Call SetPlayerGuild(n, "")
            Call PlayerMsgCombat(n, "You have been removed from the guild!", Yellow)
            End If
        Else
            Call PlayerMsgCombat(index, "That Player is not in your guild!", BrightRed)
        End If
    Else
    Call PlayerMsgCombat(index, "You do not have the rank to do this action!", BrightRed)
    End If
    Exit Sub
    End If
    
    
    ' ::::::::::::::::::::::::::
    ' :: Add character packet ::
    ' ::::::::::::::::::::::::::
    If LCase(parse(0)) = "addchar" Then
        If Not IsPlaying(index) Then
            Name = parse(1)
            Sex = Val(parse(2))
            Class = Val(parse(3))
            Race = Val(parse(4))
            charnum = Val(parse(5))
            Temp = Val(parse(6))
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Then
                Call AlertMsg(index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            If InStr(Name, " ") Then
                Call AlertMsg(index, "Your Character name can not contain spaces")
                Exit Sub
            End If
            
            ' Prevent being me
            If (LCase(Trim(Name)) = "psychoboy") Or (LCase(Trim(Name)) = "lith") Then
                Call AlertMsg(index, "Lets get one thing straight, you are not me, ok? :)")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
                                    
            ' Prevent hacking
            If charnum < 1 Or charnum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(index, "Invalid Sex (dont laugh)")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Class < 0 Or Class > Max_Classes Then
                Call HackingAttempt(index, "Invalid Class")
                Exit Sub
            End If
        
        ' Prevent hacking
            If Race < 0 Or Race > Max_Races Then
                Call HackingAttempt(index, "Invalid Race")
                Exit Sub
            End If
        
            ' Check if char already exists in slot
            If CharExist(index, charnum) Then
                Call AlertMsg(index, "Character already exists!")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If FindChar(Name) Then
                Call AlertMsg(index, "Sorry, but that name is in use!")
                Exit Sub
            End If
        
            ' Everything went ok, add the character
            Call AddChar(index, Name, Sex, Class, Race, charnum, Temp)
           ' Call AddChar(Index, Name, Sex, CLass, CharNum, tEmp)
            Call SavePlayer(index)
            'Call SavePlayer(Index)
            'Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(index, "Character has been created!")
            Call AlertMsg(index, "Character has been created!")
        End If
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::::::
    ' :: Deleting character packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "delchar" Then
        If Not IsPlaying(index) Then
            charnum = Val(parse(1))
        
            ' Prevent hacking
            If charnum < 1 Or charnum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(index, charnum)
            'Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(index, "Character has been deleted!")
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Using character packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(parse(0)) = "usechar" Then
        If Not IsPlaying(index) Then
            charnum = Val(parse(1))
        
            ' Prevent hacking
            If charnum < 1 Or charnum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(index, charnum) Then
                player(index).charnum = charnum
                Call LoadAdminCmds(index)
                Call JoinGame(index)
            
                charnum = player(index).charnum
                'Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption
                
                ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindChar(GetPlayerName(index)) Then
                    f = FreeFile
                    Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(index)
                    Close #f
                End If
            Else
                Call AlertMsg(index, "Character does not exist!")
            End If
        End If
        Exit Sub
    End If
    '::::::::::::::::::
    ':: Get the Time ::
    '::::::::::::::::::

    If LCase(parse(0)) = "time" Then
    Dim DisplaySeconds As Byte
    Dim DisplaySecondsTime As String

    DisplaySeconds = TimeSeconds / 5
        If DisplaySeconds < 10 Then
        DisplaySecondsTime = "0" & DisplaySeconds
        Else
        DisplaySecondsTime = DisplaySeconds
        End If
    Call PlayerMsgCombat(index, "The current time in Silent Shadows is " & TimeHour & ":" & DisplaySecondsTime, EmoteColor)
    Exit Sub
    End If
        
    ' :::::::::::::::::
    ' :: Bind Player ::
    ' :::::::::::::::::
    If LCase(parse(0)) = "bind" Then
    If MAP(GetPlayerMap(index)).BootMap = 1 Then
        Call SetPlayerBindMap(index, GetPlayerMap(index))
        Call SetPlayerBindX(index, GetPlayerX(index))
        Call SetPlayerBindY(index, GetPlayerY(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call PlayerMsgCombat(index, "You have been recorded in the book of life.", Yellow)
    Else
        Call PlayerMsgCombat(index, "This is not a place to record your life.", Yellow)
    End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If LCase(parse(0)) = "saymsg" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        Msg = SF_replaceAllOnce(Msg, "                            ", "-")
        'Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " says, '" & Msg & "'", SayColor)
        Exit Sub
    End If
    
    If LCase(parse(0)) = "gchat" Then
        Msg = parse(1)
        If GetPlayerGuildRank(index) <= 0 Or GetPlayerGuild(index) = "" Then
        Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Guild Text Modification")
                Exit Sub
            End If
        Next i
    s = GetPlayerName(index) & "<guild>: " & Msg
       ' Call AddLog(s, PLAYER_LOG)
        Call GlobalMsgGuild(index, s, Yellow)
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(parse(0)) = "partychat" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Guild Text Modification")
                Exit Sub
            End If
        Next i
    s = GetPlayerName(index) & "<party>: " & Msg
       ' Call AddLog(s, PLAYER_LOG)
        Call GlobalMsgParty(index, s, Yellow)
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(parse(0)) = "emotemsg" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        'Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Msg, EmoteColor)
        Exit Sub
    End If
    
    If LCase(parse(0)) = "broadcastmsg" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        
        'Msg = LCase(Msg)
       ' Call AddLog(s, PLAYER_LOG)
        'Msg = SF_replaceAllOnce(Msg, "ass", "***")
        'Msg = SF_replaceAllOnce(Msg, "fuck", "****")
        'Msg = SF_replaceAllOnce(Msg, "shit", "****")
        'Msg = SF_replaceAllOnce(Msg, "dick", "****")
        'Msg = SF_replaceAllOnce(Msg, "cunt", "****")
        Msg = SF_replaceAllOnce(Msg, "                            ", "-")
        'Msg = SF_replaceAllOnce(Msg, "whore", "*****")
        'Msg = SF_replaceAllOnce(Msg, "cock", "****")
        'Msg = SF_replaceAllOnce(Msg, "penis", "*****")
        'Msg = SF_replaceAllOnce(Msg, "bitch", "*****")
        'Msg = SF_replaceAllOnce(Msg, "slut", "****")
        'Msg = SF_replaceAllOnce(Msg, "cum", "***")
        'Msg = SF_replaceAllOnce(Msg, "skank", "*****")
        'Msg = SF_replaceAllOnce(Msg, "dildo", "*****")
        'Msg = SF_replaceAllOnce(Msg, "bastard", "*******")
        'Msg = SF_replaceAllOnce(Msg, "vagina", "******")
        'Msg = SF_replaceAllOnce(Msg, "horny", "*****")
        'Msg = SF_replaceAllOnce(Msg, "lag", "***")
        'Msg = SF_replaceAllOnce(Msg, " cl*** ", " class ")
        'Msg = SF_replaceAllOnce(Msg, " p*** ", " pass ")
        'Msg = SF_replaceAllOnce(Msg, "(1)", "XD")
        'Msg = SF_replaceAllOnce(Msg, "(2)", "=P")
        'Msg = SF_replaceAllOnce(Msg, "(3)", "=X")
        'Msg = SF_replaceAllOnce(Msg, "(4)", "=O")
        'Msg = SF_replaceAllOnce(Msg, "(5)", "=D~")
        'Msg = SF_replaceAllOnce(Msg, "(6)", "T_T")
        'Msg = SF_replaceAllOnce(Msg, "(7)", "=S")
        s = GetPlayerName(index) & " shouts, '" & Msg & "'"
        
        
        
        
        
        If GetPlayerOocSwitch(index) > 0 Then
        If CanGlobal = 1 Then
        If GetPlayerLevel(index) < 10 Then
        Call PlayerMsg(index, "You are not of the proper level to use global message yet. If you need help, contact a Shadow Knight or use the forums on the website www.silent-shadows.com", BrightRed)
        Exit Sub
        End If
        If GetTickCount >= player(index).GlobalChatTimer + 10000 Then
        player(index).GlobalChatTimer = GetTickCount
        Call GlobalMsgOoc(s, BroadcastColor)
        'If MAP(GetPlayerMap(index)).Down > 0 Then
        'Call MapMsg(MAP(GetPlayerMap(index)).Down, s, BroadcastColor)
        'End If
        'If MAP(GetPlayerMap(index)).Left > 0 Then
        'Call MapMsg(MAP(GetPlayerMap(index)).Left, s, BroadcastColor)
        'End If
        'If MAP(GetPlayerMap(index)).Right > 0 Then
        'Call MapMsg(MAP(GetPlayerMap(index)).Right, s, BroadcastColor)
        'End If
        'If MAP(GetPlayerMap(index)).Up > 0 Then
        'Call MapMsg(MAP(GetPlayerMap(index)).Up, s, BroadcastColor)
        'End If
       'Call MapMsg(GetPlayerMap(index), s, BroadcastColor)
        
        End If
        Call TextAdd(frmServer.txtText, s, True)
        Call AddLog(s, PLAYER_LOG)
        End If
        End If
        Exit Sub
    End If
    
     If LCase(parse(0)) = "tcmsg" Then
     
     'Call PlayerMsg(index, "Currently Disabled!", BrightRed)
     'Exit Sub
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        
        'Msg = LCase(Msg)
        Call AddLog(s, PLAYER_LOG)
        Msg = SF_replaceAllOnce(Msg, "                            ", "-")
        s = GetPlayerName(index) & "(Trade): " & Msg
        
        
        
        
        
        If GetPlayertcSwitch(index) > 0 Then
        If player(index).GlobalPriv = 1 Then
       
        If CanGlobal = 1 Then
If GetPlayerLevel(index) < 10 And GetPlayerCraft(index) < 15 And GetPlayerMining(index) < 15 And GetPlayerFishing(index) < 15 Then
       Call PlayerMsg(index, "Your Level or Skill is to low to use this form of communication.", BrightRed)
       Exit Sub
       End If
       
        If GetTickCount >= player(index).TradeChatTimer + 10000 Then
        player(index).TradeChatTimer = GetTickCount
        Call GlobalMsgtc(s, BrightGreen)
        Call TextAdd(frmServer.txtText, s, True)
        End If
        End If
        End If
        End If
        Exit Sub
        
    End If
    
    If LCase(parse(0)) = "globalmsg" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            s = "(global) " & GetPlayerName(index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call GlobalMsg2(s, GlobalColor)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(parse(0)) = "adminmsg" Then
        Msg = parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(index) & ") " & Msg, AdminColor)
        End If
        Exit Sub
    End If
    
    If LCase(parse(0)) = "playermsg" Then
        MsgTo = FindPlayer(parse(1))
        Msg = parse(2)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Msg & "'", TellColor)
                Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
           ' Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", Green)
        End If
        
        Exit Sub
    End If
    
    If LCase(parse(0)) = "ooc" Then
    Call SetPlayerOoc(index, parse(1))
    Exit Sub
    End If
    If LCase(parse(0)) = "tcswitch" Then
    Call SetPlayertc(index, parse(1))
    Exit Sub
    End If
    
    If LCase(parse(0)) = "emote" Then
    n = Val(parse(1))
    Call SendDataToMap(GetPlayerMap(index), "EMOTEGFX" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
    Exit Sub
    End If
    
    If LCase(parse(0)) = "guildwho" Then
    Call SendGuildWhosOnline(index)
    Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(parse(0)) = "playermove" And player(index).GettingMap = NO Then
        Dir = Val(parse(1))
        Movement = Val(parse(2))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Movement < 1 Or Movement > 2 Then
            Call HackingAttempt(index, "Invalid Movement")
            Exit Sub
        End If
        
        ' Prevent player from moving if they have casted a spell
        If player(index).CastedSpell = YES Then
            ' Check if they have already casted a spell, and if so we can't let them move
            If GetTickCount > player(index).AttackTimer + 1000 Then
                player(index).CastedSpell = NO
            Else
                Call SendPlayerXY(index)
                Exit Sub
            End If
        End If
        
        Call PlayerMove(index, Dir, Movement)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(parse(0)) = "playerdir" And player(index).GettingMap = NO Then
        Dir = Val(parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(index, Dir)
        Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    
    ' ::::::::::::::::::::::
    ' :: Sell item packet ::
    ' ::::::::::::::::::::::
    If LCase(parse(0)) = "sellitem" Then
        InvNum = Val(parse(1))
        charnum = player(index).charnum
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If charnum < 1 Or charnum > MAX_CHARS Then
            Call HackingAttempt(index, "Invalid CharNum")
            Exit Sub
        End If
        
        If MAP(GetPlayerMap(index)).Shop < 1 Then
            Call PlayerMsgCombat(index, "There is no shop here!", BrightRed)
            Exit Sub
        End If
        
        Temp = FindOpenMapItemSlot(GetPlayerMap(index))
        If Temp <> 0 Then
        
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Weapon before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Accessory before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Armor before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Shield before selling an item of this type.", BrightRed)
        Exit Sub
        End If
            If Item(GetPlayerInvItemNum(index, InvNum)).SellValue > 0 Then
            'Call SpawnItemSlot(tEmp, MapItem(GetPlayerMap(index), i).Num, Ammount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            'Call SpawnItemSlot(tEmp, 1, Item(GetPlayerInvItemNum(index, InvNum)).SellValue, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            
            Call GiveItem(index, 1, Item(GetPlayerInvItemNum(index, InvNum)).SellValue)
            Call PlayerMsgCombat(index, "The item sold for " & Item(GetPlayerInvItemNum(index, InvNum)).SellValue & " Gold!", BrightRed)
            Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 1)
            'Call SetPlayerInvItemNum(index, InvNum, 0)
            'Call SetPlayerInvItemValue(index, InvNum, 0)
            'Call SetPlayerInvItemDur(index, InvNum, 0)
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            Else
            Call PlayerMsgCombat(index, "The item was given away!", BrightRed)
            Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 1)
            Call SendInventoryUpdate(index, InvNum)
            
            End If
        Else
            Call PlayerMsgCombat(index, "Too many items on the ground to drop more gold", BrightRed)
            Exit Sub
        End If
    End If
        
        
      ' ::::::::::::
      ' :: STUCK :::
      ' ::::::::::::
      If LCase(parse(0)) = "stuck" Then
      If GetTickCount > player(index).StuckTimer + 10000 Then
      Call SendPlayerXY(index)
      Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
      Call SendPlayerXY(index)
      player(index).StuckTimer = GetTickCount
      End If
      Exit Sub
      End If
'moo
'EVIL


    ' ::::::::::::::::::
    ' :: Item Details ::
    ' ::::::::::::::::::
    
    If LCase(parse(0)) = "itemdetails" Then
    n = Val(parse(1))
    Call SendItemInvDetails(index, n)
    Exit Sub
    End If
     If LCase(parse(0)) = "shopitemdetails" Then
    n = Val(parse(1))
    Call SendItemShopDetails(index, n)
    Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "useitem" Then
        InvNum = Val(parse(1))
        charnum = player(index).charnum
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If charnum < 1 Or charnum > MAX_CHARS Then
            Call HackingAttempt(index, "Invalid CharNum")
            Exit Sub
        End If
        
        If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            
            n = Item(GetPlayerInvItemNum(index, InvNum)).Data2
            'n = Int(n / 2)
            
            'see if ur the wrong class
            If Item(GetPlayerInvItemNum(index, InvNum)).Class > 0 And Val(GetPlayerClass(index)) + 1 <> Item(GetPlayerInvItemNum(index, InvNum)).Class Then
            Call PlayerMsgCombat(index, "You are the wrong class to equip this", BrightRed)
            Exit Sub
            End If
            
            ' Find out what kind of item it is
Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                
               Case ITEM_TYPE_NONE
               
                    
                    If GetPlayerInvItemNum(index, InvNum) = 968 Then
                    If GetTickCount > player(index).FishingTimer + 5000 Then
                    player(index).FishingTimer = GetTickCount
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_FISHING Then
                    If HasItem(index, 944) Then
                    X = GetPlayerFishing(index)
                    Call TakeItem(index, 944, 1)
                    Randomize
                    y = Int(Rnd * 100) + 1
                    If y <= X Then
                    Randomize
                    y = Int(Rnd * 150) + 1
                    If y = 1 Then
                    Call PlayerMsgCombat(index, "You Have gained a level in Fishing!", BrightRed)
                    Call SetPlayerFishing(index, GetPlayerFishing(index) + 1)
                    End If
                    Call PlayerMsgCombat(index, "You Caught a Fish!", Yellow)
                    Randomize
                    X = Int(Rnd * 3) + 1
                    If X = 1 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1, 1)
                    ElseIf X = 2 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2, 1)
                    ElseIf X = 3 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3, 1)
                    End If
                    
                    
                    Else
                    Call PlayerMsgCombat(index, "You didn't catch a thing!", Yellow)
                    End If
                    Else
                    Call PlayerMsgCombat(index, "You have no fish Bait!", Yellow)
                    End If
                    End If
                    End If
                    End If
                    
                    If GetPlayerInvItemNum(index, InvNum) = 967 Then
                    If GetTickCount > player(index).MiningTimer + 5000 Then
                    If TempTile(GetPlayerMap(index)).MineralCount(GetPlayerX(index), GetPlayerY(index)) < 5 Then
                    TempTile(GetPlayerMap(index)).MineralCount(GetPlayerX(index), GetPlayerY(index)) = TempTile(GetPlayerMap(index)).MineralCount(GetPlayerX(index), GetPlayerY(index)) + 1
                    TempTile(GetPlayerMap(index)).MineralTimer(GetPlayerX(index), GetPlayerY(index)) = GetTickCount
                    Else
                    If TempTile(GetPlayerMap(index)).MineralTimer(GetPlayerX(index), GetPlayerY(index)) + 3600000 < GetTickCount Then
                    'do other stuff
                    Else
                    Call PlayerMsg(index, "There are no more minerals to mine here!", BrightRed)
                    Exit Sub
                    End If
                    End If
                    player(index).MiningTimer = GetTickCount
                    If MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_MINING Then
                    X = GetPlayerMining(index)
                    
                    y = Int(Rnd * 10)
                    If y = 2 Then
                    Call PlayerMsgCombat(index, "Your Mining Pick Broke!", BrightRed)
                    Call TakeItem(index, 967, 1)
                    Exit Sub
                    End If
                    
                    If X < 0 Then X = 0
                    Randomize
                    y = Int(Rnd * 110)
                    
                    If y <= X Then
                    Randomize
                     y = Int(Rnd * 150) + 1
                    If y = 1 Then
                    Call PlayerMsgCombat(index, "You Have gained a level in Mining!", BrightRed)
                    Call SetPlayerMining(index, GetPlayerMining(index) + 1)
                    End If
                    Randomize
                    X = Int(Rnd * 3) + 1
                     Call PlayerMsgCombat(index, "You have mined some sort of Mineral!", Yellow)
                    
                    If X = 1 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1, 1)
                    ElseIf X = 2 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2, 1)
                    ElseIf X = 3 Then
                    Call GiveItem(index, MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3, 1)
                    End If
                    
                    
                   Else
                   Call PlayerMsgCombat(index, "You didn't mine anything!", Yellow)
                   End If
                   End If
                   End If
                    
                    End If
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(index) Then
                     If Int(GetPlayerDEF(index)) < n Then
                      Call PlayerMsgCombat(index, "Your defense is to low to wear this armor!  Required DEF (" & n & ")", BrightRed)
                      Exit Sub
                     End If
                     If GetPlayerArmorSlot(index) > 0 Then
                      Temp = GetPlayerArmorSlot(index)
                      Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                      Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                      Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                     End If
                      Call SetPlayerArmorSlot(index, InvNum)
                      Call SetPlayerSTR(index, (GetPlayerSTR(index) + Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                      Call SetPlayerDEF(index, (GetPlayerDEF(index) + Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                      Call SetPlayerMAGI(index, (GetPlayerMAGI(index) + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                      If Item(GetPlayerInvItemNum(index, InvNum)).DEFmod + Item(GetPlayerInvItemNum(index, InvNum)).STRmod + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod > 0 Then
                        Call PlayerMsgCombat(index, "This Armor is enchanted...", Cyan)
                      End If
                       If GetPlayerSex(index) = 0 Then
                        Call SetPlayerSprite(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE)
                       Else
                        Call SetPlayerSprite(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE + 1)
                       End If
                       SendPlayerData (index)
                      Else
                        Call SetPlayerArmorSlot(index, 0)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                        If GetPlayerSex(index) = 0 Then
                        Call SetPlayerSprite(index, 0)
                        Else
                        Call SetPlayerSprite(index, 1)
                        End If
                        SendPlayerData (index)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(index) Then
                        If Int(GetPlayerSTR(index)) < n Then
                            Call PlayerMsgCombat(index, "Your strength is to low to hold this weapon!  Required STR (" & n & ")", BrightRed)
                            Exit Sub
                        End If
                        
                        If GetPlayerWeaponSlot(index) > 0 Then
                         Temp = GetPlayerWeaponSlot(index)
                         Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                         Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                         Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        End If
                        
                        Call SetPlayerWeaponSlot(index, InvNum)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) + Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) + Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                            If Item(GetPlayerInvItemNum(index, InvNum)).DEFmod + Item(GetPlayerInvItemNum(index, InvNum)).STRmod + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod > 0 Then
                            Call PlayerMsgCombat(index, "This Weapon is enchanted...", Cyan)
                            End If
                        Call SetPlayerSprite4(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE)
                        Call SendPlayerData(index)
                        Else
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                        Call SetPlayerSprite4(index, 1000)
                        Call SendPlayerData(index)
                    End If
                    Call SendWornEquipment(index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(index) Then
                        If Int(GetPlayerSPEED(index)) < n Then
                            Call PlayerMsgCombat(index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n & ")", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerHelmetSlot(index) > 0 Then
                        Temp = GetPlayerHelmetSlot(index)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                        End If
                        Call SetPlayerHelmetSlot(index, InvNum)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) + Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) + Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                        'Call SetPlayerSprite3(Index, Item(GetPlayerInvItemNum(Index, InvNum)).SPRITE)
                    Else
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                        Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                        Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                       ' Call SetPlayerSprite3(Index, 1000)
                    End If
                    Call SendWornEquipment(index)
                    Call SendPlayerData(index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(index) Then
                     If Int(GetPlayerDEF(index)) < n Then
                      Call PlayerMsgCombat(index, "Your Defense is to low to wield this Shield!  Required Defense (" & n & ")", BrightRed)
                      Exit Sub
                     End If
                     If GetPlayerShieldSlot(index) > 0 Then
                      Temp = GetPlayerShieldSlot(index)
                      Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, Temp)).STRmod))
                      Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, Temp)).DEFmod))
                      Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, Temp)).MAGImod))
                      Call SetPlayerShieldSlot(index, 0)
                     End If
                     If Int(GetPlayerDEF(index)) < n Then
                      Call PlayerMsgCombat(index, "Your Defense is to low to wield this Shield!  Required Defense (" & n & ")", BrightRed)
                      Call SetPlayerSprite3(index, 4000)
                      Call SendWornEquipment(index)
                      Call SendPlayerData(index)
                      Exit Sub
                     End If
                      Call SetPlayerShieldSlot(index, InvNum)
                      Call SetPlayerSTR(index, (GetPlayerSTR(index) + Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                      Call SetPlayerDEF(index, (GetPlayerDEF(index) + Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                      Call SetPlayerMAGI(index, (GetPlayerMAGI(index) + Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                      Call SetPlayerSprite3(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE)
                    Else
                     Call SetPlayerShieldSlot(index, 0)
                     Call SetPlayerSTR(index, (GetPlayerSTR(index) - Item(GetPlayerInvItemNum(index, InvNum)).STRmod))
                     Call SetPlayerDEF(index, (GetPlayerDEF(index) - Item(GetPlayerInvItemNum(index, InvNum)).DEFmod))
                     Call SetPlayerMAGI(index, (GetPlayerMAGI(index) - Item(GetPlayerInvItemNum(index, InvNum)).MAGImod))
                     Call SetPlayerSprite3(index, 1000)
                    End If
                    Call SendWornEquipment(index)
                    Call SendPlayerData(index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(index, GetPlayerHP(index) + Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(index, GetPlayerMP(index) + Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(index, GetPlayerSP(index) + Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendSP(index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(index, GetPlayerHP(index) - Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(index, GetPlayerMP(index) - Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(index, GetPlayerSP(index) - Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendSP(index)
                    
                Case ITEM_TYPE_PETEGG
                    Call SetPet(index, Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    Call SendPlayerData(index)
                    
                'for the map edit items
                
                Case ITEM_TYPE_MASKEDIT
                    If Item(player(index).Char(charnum).Inv(InvNum).Num).Data2 = 0 Then
                        If MAP(GetPlayerMap(index)).Moral = MAP_MORAL_HOUSE Then
                            MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Mask = Item(player(index).Char(charnum).Inv(InvNum).Num).Data1
                        End If
                    Else
                        If MAP(GetPlayerMap(index)).Moral = MAP_MORAL_HOUSE Then
                            MAP(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Fringe = Item(player(index).Char(charnum).Inv(InvNum).Num).Data1
                        End If
                    End If
                    MAP(GetPlayerMap(index)).Revision = MAP(GetPlayerMap(index)).Revision + 1
                    ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                Call PlayerWarp(i, GetPlayerMap(index), GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
        SaveMap (GetPlayerMap(index))
                    
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(index)
                        Case DIR_UP
                            If GetPlayerY(index) > 0 Then
                                X = GetPlayerX(index)
                                y = GetPlayerY(index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(index) < MAX_MAPY Then
                                X = GetPlayerX(index)
                                y = GetPlayerY(index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(index) > 0 Then
                                X = GetPlayerX(index) - 1
                                y = GetPlayerY(index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(index) < MAX_MAPY Then
                                X = GetPlayerX(index) + 1
                                y = GetPlayerY(index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If MAP(GetPlayerMap(index)).Tile(X, y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(index, InvNum) = MAP(GetPlayerMap(index)).Tile(X, y).Data1 Then
                            TempTile(GetPlayerMap(index)).DoorOpen(X, y) = YES
                            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                            
                            ' Check if we are supposed to take away the item
                            If MAP(GetPlayerMap(index)).Tile(X, y).Data2 = 1 Then
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsgCombat(index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If spell(n).ClassReq - 1 = GetPlayerClass(index) Or spell(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            If spell(n).Type <> SPELL_TYPE_CRAFT Then
                            i = GetSpellReqLevel(index, n)
                            Else
                            i = 1
                            End If
                            If i <= GetPlayerLevel(index) Then
                                i = FindOpenSpellSlot(index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(index, n) Then
                                        Call SetPlayerSpell(index, i, n)
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsgCombat(index, "You study the spell carefully...", Yellow)
                                        Call PlayerMsgCombat(index, "You have learned a new spell!", White)
                                    Else
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsgCombat(index, "You have already learned this spell!  The spells crumbles into dust.", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsgCombat(index, "You have learned all that you can learn!", BrightRed)
                                End If
                            Else
                                Call PlayerMsgCombat(index, "You must be level " & i & " to learn this spell.", White)
                            End If
                        Else
                            Call PlayerMsgCombat(index, "This spell can only be learned by a " & GetClassName(spell(n).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsgCombat(index, "This scroll is not connected to a spell, please inform an admin!", White)
                    End If
                Case ITEM_TYPE_SCHANGE
                Temp = GetPlayerSex(index)
                If Temp = 0 Then
                    Call SetPlayerSprite3(index, Item(player(index).Char(charnum).Inv(InvNum).Num).Data1)
                Else
                    Call SetPlayerSprite3(index, Val(Item(player(index).Char(charnum).Inv(InvNum).Num).Data1))
                End If
                    Call SendPlayerData(index)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                
                Case ITEM_TYPE_BOOK
                Tmpstr = "ITEMBOOK" & SEP_CHAR & Item(player(index).Char(charnum).Inv(InvNum).Num).Data2 & SEP_CHAR & Item(player(index).Char(charnum).Inv(InvNum).Num).Name & SEP_CHAR & END_CHAR
                    Call SendDataTo(index, Tmpstr)
                    
                 Case ITEM_TYPE_BARBER
                    If GetPlayerSex(index) = 0 Then
                    Call SetPlayerSprite2(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE)
                    Else
                    Call SetPlayerSprite2(index, Item(GetPlayerInvItemNum(index, InvNum)).SPRITE + 12)
                    End If
                    Call SendPlayerData(index)
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
                
                    
                Case ITEM_TYPE_RANDITEM
                Dim randtempy As Double
randtempy = (Rnd * 1)
If randtempy > 0 And randtempy < 0.33334 Then Call GiveItem(index, Val(Item(player(index).Char(charnum).Inv(InvNum).Num).Data1), 0)
If randtempy > 0.33333 And randtempy < 0.66667 Then Call GiveItem(index, Val(Item(player(index).Char(charnum).Inv(InvNum).Num).Data2), 0)
If randtempy > 0.6666 Then Call GiveItem(index, Val(Item(player(index).Char(charnum).Inv(InvNum).Num).Data3), 0)
Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).Num, 0)
Call GiveItem(index, Val(Item(player(index).Char(charnum).Inv(InvNum).Num).Data1), 0)
                    
            End Select
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If LCase(parse(0)) = "attack" Then
Dim PlayerSP As Long
        ' Try to attack a player
        For i = 1 To MAX_PLAYERS
           ' Call TextAdd(frmServer.txtText, "Attack Player", True)
            ' Make sure we dont try to attack ourselves
            If i <> index Then
                ' Can we attack the player?
                If CanAttackPlayer(index, i) Then
                If GetPlayerWeaponSlot(index) > 0 Then
    PlayerSP = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data2 / 14
    Else
    PlayerSP = 2
    End If
   If PlayerSP < 2 Then PlayerSP = 2
    PlayerSP = GetPlayerSP(index) - (PlayerSP)
    '
    Call SetPlayerSP(index, PlayerSP)
    Call SendSP(index)
    If GetPlayerSP(index) <= 0 Then
    Call PlayerMsgCombat(index, "You are to tired to fight.", Blue)
    Exit Sub
    End If
                'Call TextAdd(frmServer.txtText, "can Attack Player", True)
                    If Not CanPlayerBlockHit(i) Then
                   ' Call TextAdd(frmServer.txtText, "can blackAttack Player", True)
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsgCombat(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsgCombat(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage, "z")
                        Else
                            Call PlayerMsgCombat(index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsgCombat(index, GetPlayerName(i) & "'s " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsgCombat(i, "Your " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next i
        
        ' Try to attack a npc
        For i = 1 To MAX_MAP_NPCS
            
            ' Can we attack the npc?
            If CanAttackNpc(index, i) Then
            'Call TextAdd(frmServer.txtText, "CanAttack NPC", True)
                ' Get the damage we can do
                If GetPlayerWeaponSlot(index) > 0 Then
                PlayerSP = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data2 / 14
        Else
        PlayerSP = 2
        End If
   If PlayerSP < 2 Then PlayerSP = 2
    PlayerSP = GetPlayerSP(index) - (PlayerSP)
    Call SetPlayerSP(index, PlayerSP)
    Call SendSP(index)
    If GetPlayerSP(index) <= 0 Then
    Call PlayerMsgCombat(index, "You are to tired to fight.", Blue)
    Exit Sub
    End If
                If Not CanPlayerCriticalHit(index) Then
               ' Call TextAdd(frmServer.txtText, "not critical Attack NPC", True)
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).Num).DEF / 2)
                Else
               ' Call TextAdd(frmServer.txtText, "Critical Attack NPC", True)
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).Num).DEF / 2)
                    Call PlayerMsgCombat(index, "You feel a surge of energy upon swinging!", BrightCyan)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage, "z")
                  '  Call TextAdd(frmServer.txtText, "No DMG Attack NPC", True)
                Else
                'Call TextAdd(frmServer.txtText, "0 attack Attack NPC", True)
                    Call PlayerMsgCombat(index, "Your attack does nothing.", BrightRed)
                End If
                Exit Sub
            End If
        Next i
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Use stats packet ::
    ' ::::::::::::::::::::::
    If LCase(parse(0)) = "usestatpoint" Then
        PointType = Val(parse(1))
        
        ' Prevent hacking
        If (PointType < 0) Or (PointType > 6) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(index) > 0 Then
            ' Take away a stat point
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
            
            ' Everything is ok
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more strength!", White)
                Case 1
                    Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more defense!", White)
                Case 2
                    Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more magical abilities!", White)
                Case 3
                    Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more speed!", White)
                Case 4
                    Call SetPlayerCraft(index, GetPlayerCraft(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more Crafting experiance!", White)
                Case 5
                    Call SetPlayerMining(index, GetPlayerMining(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more Mining experiance!", White)
                Case 6
                    Call SetPlayerFishing(index, GetPlayerFishing(index) + 1)
                    Call PlayerMsgCombat(index, "You have gained more Fishing Experiance!", White)
                
                    End Select
        Else
            Call PlayerMsgCombat(index, "You have no skill points to train with!", BrightRed)
        End If
        
        ' Send the update
        Call SendStats(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "playerinforequest" Then
     Name = parse(1)
     
     i = FindPlayer(Name)
     If i > 0 Then
      Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
      Call PlayerMsgCombat(index, "Class: " & GetClassName(GetPlayerClass(i)) & "  Race: " & GetRaceName(GetPlayerRace(i)), BrightGreen)
      If GetPlayerAccess(index) > 0 Then
       Call PlayerMsgCombat(index, "Account: " & Trim(player(i).Login), BrightGreen)
       Call PlayerMsgCombat(index, "Access: " & Trim(player(i).Char(player(i).charnum).Access), BrightGreen)
       Call PlayerMsgCombat(index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
       Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
       Call PlayerMsgCombat(index, "STR: " & GetPlayerSTR(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  Speed: " & GetPlayerSPEED(i), BrightGreen)
       Tmpstr = GetPlayerIP(i)
       Call PlayerMsgCombat(index, "IP:" & Tmpstr, BrightGreen)
       Temp = GetPlayerMap(i)
       Call PlayerMsgCombat(index, "Map#:" & Temp, BrightGreen)
       n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
       i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
       If n > 100 Then n = 100
       If i > 100 Then i = 100
       Call PlayerMsgCombat(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
       End If
      Else
       Call PlayerMsgCombat(index, "Player is not online.", White)
      End If
        Exit Sub
    End If
       
    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
                Call PlayerMsgCombat(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
                Call PlayerMsgCombat(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
                Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot warp to yourself!", White)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Warp to me packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "warptome" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMsgCombat(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
                Call PlayerMsgCombat(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
                Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot warp yourself to yourself!", White)
        End If
        
        Exit Sub
    End If



    ' ::::::::::::::::::::::::
    ' :: Warp to map packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "warpto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The map
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
        Call PlayerMsgCombat(index, "You have been warped to map #" & n, BrightBlue)
        Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "setsprite" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        n = Val(parse(1))
        
        Call SetPlayerSprite(index, n)
        Call SendPlayerData(index)
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(parse(0)) = "getstats" Then
        Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(index) & " -=-", White)
        percentage = (GetPlayerExp(index) / GetPlayerNextLevel(index)) * 100
        Call PlayerMsgCombat(index, "Level: " & GetPlayerLevel(index) & "  Exp: " & percentage & "%", White)
        Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index) & "  MP: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index) & "  SP: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index), White)
        Call PlayerMsgCombat(index, "STR: " & GetPlayerSTR(index) & "  DEF: " & GetPlayerDEF(index) & "  MAGI: " & GetPlayerMAGI(index) & "  Speed: " & GetPlayerSPEED(index), BrightGreen)
        Call PlayerMsgCombat(index, "Crafting: " & GetPlayerCraft(index) & "  Fishing: " & GetPlayerFishing(index) & "  Mining: " & GetPlayerMining(index) & "  Avail Points: " & GetPlayerPOINTS(index), BrightGreen)
        n = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
        i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        If n > 100 Then n = 100
        If i > 100 Then i = 100
        Call PlayerMsgCombat(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
        Exit Sub
    End If
    
        
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requestnewmap" Then
        Dir = Val(parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(index, Dir, 1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "mapdata" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        MapNum = GetPlayerMap(index)
        MAP(MapNum).Name = parse(n + 1)
        MAP(MapNum).Revision = MAP(MapNum).Revision + 1
        MAP(MapNum).Moral = Val(parse(n + 3))
        MAP(MapNum).Up = Val(parse(n + 4))
        MAP(MapNum).Down = Val(parse(n + 5))
        MAP(MapNum).Left = Val(parse(n + 6))
        MAP(MapNum).Right = Val(parse(n + 7))
        MAP(MapNum).Music = Val(parse(n + 8))
        MAP(MapNum).BootMap = Val(parse(n + 9))
        MAP(MapNum).BootX = Val(parse(n + 10))
        MAP(MapNum).BootY = Val(parse(n + 11))
        MAP(MapNum).Shop = Val(parse(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                MAP(MapNum).Tile(X, y).Ground = Val(parse(n))
                MAP(MapNum).Tile(X, y).Mask = Val(parse(n + 1))
                MAP(MapNum).Tile(X, y).Anim = Val(parse(n + 2))
                MAP(MapNum).Tile(X, y).Mask2 = Val(parse(n + 3))
                MAP(MapNum).Tile(X, y).M2Anim = Val(parse(n + 4))
                MAP(MapNum).Tile(X, y).Fringe = Val(parse(n + 5))
                MAP(MapNum).Tile(X, y).FAnim = Val(parse(n + 6))
                MAP(MapNum).Tile(X, y).Fringe2 = Val(parse(n + 7))
                MAP(MapNum).Tile(X, y).F2Anim = Val(parse(n + 8))
                MAP(MapNum).Tile(X, y).Type = Val(parse(n + 9))
                MAP(MapNum).Tile(X, y).Data1 = parse(n + 10)
                MAP(MapNum).Tile(X, y).Data2 = parse(n + 11)
                MAP(MapNum).Tile(X, y).Data3 = parse(n + 12)
                
                n = n + 13
            Next X
        Next y
        
        For X = 1 To MAX_MAP_NPCS
            MAP(MapNum).Npc(X) = Val(parse(n))
            n = n + 1
            Call ClearMapNpc(X, MapNum)
        Next X
        Call SendMapNpcsToMap(MapNum)
        Call SpawnMapNpcs(MapNum)
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
        
        Exit Sub
    End If



    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(parse(0)) = "needmap" Then
        ' Get yes/no value
        s = LCase(parse(1))
                
        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            player(index).Target = 0
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            player(index).Target = 0
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "mapgetitem" Then
        Call PlayerMapGetItem(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "mapdropitem" Then
        InvNum = Val(parse(1))
        Ammount = Val(parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(index, InvNum) Then
            Call HackingAttempt(index, "Item ammount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
        
        If Item(GetPlayerInvItemNum(index, InvNum)).NoDrop > 0 Then
            Call PlayerMsgCombat(index, "This is a No Drop Item", BrightRed)
            Exit Sub
        End If
            
        Call PlayerMapDropItem(index, InvNum, Ammount)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "maprespawn" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        Call PlayerMsgCombat(index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "mapreport" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        
        For i = 1 To MAX_MAPS
            If Trim(MAP(i).Name) = "" Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim(STR(tMapStart)) & "-" & Trim(STR(tMapEnd - 1)) & ", "
                End If
                tMapStart = i + 1
                tMapEnd = i + 1
            End If
        Next i
        
        s = s & Trim(STR(tMapStart)) & "-" & Trim(STR(tMapEnd - 1)) & ", "
        s = Mid(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsgCombat(index, s, Brown)
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
    ' :: Mute player packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "muteplayer" Then
    
    ' The player index
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(index) > 0 Then
                If GetPlayerAccess(n) > 0 Then
                 Call PlayerMsgCombat(index, "You cannot mute a fellow admin", White)
                 Exit Sub
                End If
                    Call GlobalMsgCombat(GetPlayerName(n) & " has been muted by " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " has muted " & GetPlayerName(n) & ".", ADMIN_LOG)
                    'Call AlertMsg(n, "You have been muted by " & GetPlayerName(index) & "!")
                    Call SendDataTo(n, "mute" & SEP_CHAR & END_CHAR)
                    player(n).GlobalPriv = 0
                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot mute yourself!", White)
        End If
        
    Exit Sub
    End If


   

    ' :::::::::::::::::::::::::
    ' :: UnMuteplayer packet ::
    ' :::::::::::::::::::::::::
    If LCase(parse(0)) = "unmuteplayer" Then
    n = FindPlayer(parse(1))
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call GlobalMsgCombat(GetPlayerName(n) & " has been unmuted by " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " has unmuted " & GetPlayerName(n) & ".", ADMIN_LOG)
                    'Call AlertMsg(n, "You have been unmuted by " & GetPlayerName(index) & "!")
                    Call SendDataTo(n, "unmute" & SEP_CHAR & END_CHAR)
                    player(n).GlobalPriv = 1
                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot unmute yourself!", White)
        End If
    Exit Sub
    End If
       
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "kickplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
               If GetPlayerAccess(n) > 0 Then
                 Call PlayerMsgCombat(index, "You cannot kick a fellow admin", White)
                 Exit Sub
                End If
                    Call GlobalMsgCombat(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot kick yourself!", White)
        End If
                
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
    ' :: Dumb player packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "dumbplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call SendDataTo(n, "dumbass" & SEP_CHAR & END_CHAR)
                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot stupify yourself!", White)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "banlist" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        f = FreeFile
        Open App.Path & "\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            
            Call PlayerMsgCombat(index, n & ": Banned IP " & s & " by " & Name, White)
            n = n + 1
        Loop
        Close #f
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Ban destroy packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "bandestroy" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\banlist.txt")
        Call PlayerMsgCombat(index, "Ban list destroyed.", White)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "banplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                If GetPlayerAccess(n) > 0 Then
                 Call PlayerMsgCombat(index, "You cannot ban a fellow admin", White)
                 Exit Sub
                End If
                    Call BanIndex(n, index)
                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot ban yourself!", White)
        End If
                
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: PBan player packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "permban" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call GlobalMsgCombat(GetPlayerName(n) & " has been PERMENENTLY banned from " & GAME_NAME & " by " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " has banned " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Dim Packet As String
                    Packet = "BANNEDMOFO" & SEP_CHAR & END_CHAR
                    Call SendDataTo(n, Packet)
                    Call SendDataTo(n, Packet)
                    Call AlertMsg(n, "You have been PERMENENTLY banned by " & GetPlayerName(index) & "!")

                Else
                    Call PlayerMsgCombat(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "You cannot Pban yourself!", White)
        End If
                
        Exit Sub
    End If
    
    
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requesteditmap" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requestedititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(parse(0)) = "edititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The item #
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    If LCase(parse(0)) = "saveitem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
        ' Update the item
        Item(n).Name = parse(2)
        Item(n).Pic = Val(parse(3))
        Item(n).Type = Val(parse(4))
        Item(n).Data1 = Val(parse(5))
        Item(n).Data2 = (parse(6))
        Item(n).Data3 = Val(parse(7))
        Item(n).Data4 = Val(parse(8))
        Item(n).Class = Val(parse(9))
        Item(n).STRmod = Val(parse(10))
        Item(n).DEFmod = Val(parse(11))
        Item(n).MAGImod = Val(parse(12))
        Item(n).SPRITE = Val(parse(13))
        Item(n).SellValue = Val(parse(14))
        Item(n).NoDrop = Val(parse(15))
        
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requesteditnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "editnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The npc #
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(index, n)
    End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "savenpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
        ' Update the npc
        Npc(n).Name = parse(2)
        Npc(n).AttackSay = parse(3)
        Npc(n).SPRITE = Val(parse(4))
        Npc(n).SpawnSecs = Val(parse(5))
        Npc(n).Behavior = Val(parse(6))
        Npc(n).Range = Val(parse(7))
        Npc(n).DropChance = Val(parse(8))
        Npc(n).DropItem = Val(parse(9))
        Npc(n).DropItemValue = Val(parse(10))
        Npc(n).STR = Val(parse(11))
        Npc(n).DEF = Val(parse(12))
        Npc(n).SPeed = Val(parse(13))
        Npc(n).MAGI = Val(parse(14))
        Npc(n).DropChance2 = Val(parse(15))
        Npc(n).DropItem2 = Val(parse(16))
        Npc(n).DropItemValue2 = Val(parse(17))
        Npc(n).SPRITE2 = Val(parse(18))
        Npc(n).SPRITE3 = Val(parse(19))
        Npc(n).SPRITE4 = Val(parse(20))
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requesteditshop" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    If LCase(parse(0)) = "editshop" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The shop #
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    If (LCase(parse(0)) = "saveshop") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(parse(1))
        
        ' Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
        ' Update the shop
        Shop(ShopNum).Name = parse(2)
        Shop(ShopNum).JoinSay = parse(3)
        Shop(ShopNum).LeaveSay = parse(4)
        Shop(ShopNum).FixesItems = Val(parse(5))
        Shop(ShopNum).OneSale = Val(parse(6))
        
        n = 7
        For i = 1 To MAX_TRADES
            Shop(ShopNum).TradeItem(i).GiveItem = Val(parse(n))
            Shop(ShopNum).TradeItem(i).GiveValue = Val(parse(n + 1))
            Shop(ShopNum).TradeItem(i).GetItem = Val(parse(n + 2))
            Shop(ShopNum).TradeItem(i).GetValue = Val(parse(n + 3))
            n = n + 4
        Next i
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "requesteditspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "editspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(parse(0)) = "savespell") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        spell(n).Name = parse(2)
        spell(n).ClassReq = Val(parse(3))
        spell(n).LevelReq = Val(parse(4))
        spell(n).Type = Val(parse(5))
        spell(n).Data1 = Val(parse(6))
        spell(n).Data2 = Val(parse(7))
        spell(n).Data3 = Val(parse(8))
        spell(n).MPused = Val(parse(9))
        spell(n).Sfx = parse(10)
        spell(n).Gfx = parse(11)
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
   ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "setaccess" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(parse(1))
        ' The access
        i = Val(parse(2))
        
        
        ' Check for invalid access level
        If i >= 0 Or i <= 3 Then
            ' Check if player is on
            If n > 0 Then
                If GetPlayerAccess(n) <= 0 Then
                    Call GlobalMsgCombat(GetPlayerName(n) & " has been Honored with joining the Shadow Knights.", BrightBlue)
                End If
                
                Call SetPlayerAccess(n, i)
                Call SendPlayerData(n)
                Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsgCombat(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsgCombat(index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "whosonline" Then
        Call SendWhosOnline(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Set MOTD packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "setmotd" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", parse(1))
        Call GlobalMsgCombat("MOTD changed to: " & parse(1), BrightCyan)
        Call AddLog(GetPlayerName(index) & " changed MOTD to: " & parse(1), ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If LCase(parse(0)) = "trade" Then
        If MAP(GetPlayerMap(index)).Shop > 0 Then
            Call SendTrade(index, MAP(GetPlayerMap(index)).Shop)
        Else
            Call PlayerMsgCombat(index, "There is no shop here.", BrightRed)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(parse(0)) = "traderequest" Then
        ' Trade num
        n = Val(parse(1))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = MAP(GetPlayerMap(index)).Shop
        
        ' Check if inv full
        X = FindOpenInvSlot(index, Shop(i).TradeItem(n).GetItem)
        If X = 0 Then
            Call PlayerMsgCombat(index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
        y = Shop(i).TradeItem(n).GiveValue
        If Shop(i).TradeItem(n).GiveItem = 1 Then
        If MAP(GetPlayerMap(index)).BootMap - 2 = GetPlayerFaction(index) Then
        y = y * 0.9
        End If
        End If
        
            Call TakeItem(index, Shop(i).TradeItem(n).GiveItem, y)
            Call GiveItem(index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call PlayerMsgCombat(index, "The trade was successful!", Yellow)
            If Shop(i).OneSale = 1 Then
                Shop(i).TradeItem(n).GiveItem = 0
                Shop(i).TradeItem(n).GetItem = 0
                Shop(i).TradeItem(n).GiveValue = 0
                Shop(i).TradeItem(n).GetValue = 0
                Call SendUpdateShopToAll(i)
                Call SaveShop(i)
            End If
        Else
            Call PlayerMsgCombat(index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Fix item packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "fixitem" Then
        ' Inv num
        n = Val(parse(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_SHIELD Then
            Call PlayerMsg(index, "You can only fix Weapons and Armors", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(index, GetPlayerInvItemNum(index, n)) <= 0 Then
            Call PlayerMsg(index, "You have no inventory space left!", BrightRed)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(index, n)
        i = Int(Item(GetPlayerInvItemNum(index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, n)
        GoldNeeded = Int(DurNeeded * i / 5)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(index, "This item is in perfect condition!", White)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(index, 1) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(index, 1) >= GoldNeeded Then
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, Item(ItemNum).Data1)
                Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " Creditz", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, GetPlayerInvItemDur(index, n) + DurNeeded)
                Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
            End If
        Else
            Call PlayerMsg(index, "Insufficient gold to fix this item!", BrightRed)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Search packet ::
    ' :::::::::::::::::::
    If LCase(parse(0)) = "search" Then
        X = Val(parse(1))
        y = Val(parse(2))
        
        ' Prevent subscript out of range
        If X < 0 Or X > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = X And GetPlayerY(i) = y Then
                If GetPlayerAnonymous(i) < 1 Then
                ' Consider the player
                If GetPlayerLevel(i) >= GetPlayerLevel(index) + 5 Then
                    Call PlayerMsgCombat(index, "You wouldn't stand a chance.", BrightRed)
                Else
                    If GetPlayerLevel(i) > GetPlayerLevel(index) Then
                        Call PlayerMsgCombat(index, "This one seems to have an advantage over you.", Yellow)
                    Else
                        If GetPlayerLevel(i) = GetPlayerLevel(index) Then
                            Call PlayerMsgCombat(index, "This would be an even fight.", White)
                        Else
                            If GetPlayerLevel(index) >= GetPlayerLevel(i) + 5 Then
                                Call PlayerMsgCombat(index, "You could slaughter that player.", BrightBlue)
                            Else
                                If GetPlayerLevel(index) > GetPlayerLevel(i) Then
                                    Call PlayerMsgCombat(index, "You would have an advantage over that player.", Yellow)
                                End If
                            End If
                        End If
                    End If
                End If
            Else
            Call PlayerMsgCombat(index, "Your eyes cannot see this characters skills.", Yellow)
            End If
                ' Change target
                player(index).Target = i
                player(index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsgCombat(index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
               ' Call PlayerMsgCombat(Index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                Exit Sub
            End If
        Next i
        
         For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index), i).Num > 0 Then
                If MapNpc(GetPlayerMap(index), i).X = X And MapNpc(GetPlayerMap(index), i).y = y Then
                    ' Change target
                    player(index).Target = i
                    player(index).TargetType = TARGET_TYPE_NPC
                    'Call PlayerMsg(Index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Call PlayerMsgCombat(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).Num).Name) & ".", Yellow)
                    
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).Num > 0 Then
                If MapItem(GetPlayerMap(index), i).X = X And MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsgCombat(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
       
        
        Exit Sub
    End If
    
    If LCase(parse(0)) = "pinfo" Then
     If player(index).Party > 0 Then
     n = player(index).Party
        If Parties(n).Player1 > 0 Then
        i = Parties(n).Player1
        Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
        Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
        End If
        If Parties(n).Player2 > 0 Then
        i = Parties(n).Player2
        Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
        Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
        End If
        If Parties(n).Player3 > 0 Then
        i = Parties(n).Player3
        Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
        Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
        End If
        If Parties(n).Player4 > 0 Then
        i = Parties(n).Player4
        Call PlayerMsgCombat(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
        Call PlayerMsgCombat(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
        End If
        End If
    Exit Sub
    End If
    
    If LCase(parse(0)) = "pwho" Then
    If player(index).Party > 0 Then
     n = player(index).Party
     If Parties(n).Player1 > 0 Then
        i = Parties(n).Player1
        Call PlayerMsgCombat(index, GetPlayerName(i) & " is in your party.", BrightGreen)
        End If
    If Parties(n).Player2 > 0 Then
        i = Parties(n).Player2
        Call PlayerMsgCombat(index, GetPlayerName(i) & " is in your party.", BrightGreen)
        End If
    If Parties(n).Player3 > 0 Then
        i = Parties(n).Player3
        Call PlayerMsgCombat(index, GetPlayerName(i) & " is in your party.", BrightGreen)
        End If
    If Parties(n).Player4 > 0 Then
        i = Parties(n).Player4
        Call PlayerMsgCombat(index, GetPlayerName(i) & " is in your party.", BrightGreen)
        End If
        End If
        Exit Sub
        End If
        
        
        
     

    ' ::::::::::::::::::
    ' :: Party packet ::
    ' ::::::::::::::::::
    If LCase(parse(0)) = "party" Then
        n = FindPlayer(parse(1))
        
        ' Prevent partying with self
        If n = index Then
            Exit Sub
        End If
        
        If player(index).Party > 0 Then
           If Parties(player(index).Party).PartyLeader <> index Then
           Call PlayerMsgCombat(index, "You are already in a party! and not the leader!", Pink)
           Exit Sub
           End If
         If Parties(player(index).Party).NumParty >= 4 Then
        Call PlayerMsgCombat(index, "PARTY FULL!", BrightRed)
        Exit Sub
        End If
        End If
       If player(n).InParty = YES Then
       Call PlayerMsgCombat(index, "Player already in party!", BrightRed)
       Exit Sub
       End If
       
       
            
                
        ' Check for a previous party and if so drop it
        'If player(Index).InParty = YES Then
        '    If player(Index).PartyStarter = NO Then
        '    Call PlayerMsg(Index, "You are already in a party! and not the leader", Pink)
        '    Exit Sub
        '    End If
        'End If
        
        'If player(Index).PartyNum >= 3 Then
        'Call PlayerMsg(Index, "PARTY FULL!", BrightRed)
        'Exit Sub
        'End If
        
        
        If n > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsgCombat(index, "You can't join a party, you are an admin!", BrightBlue)
                Exit Sub
            End If
        
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsgCombat(index, "Admins cannot join parties!", BrightBlue)
                Exit Sub
            End If
        '
            ' Make sure they are in right level range
            If GetPlayerLevel(index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsgCombat(index, "There is more then a 5 level gap between you two, party failed.", Pink)
                Exit Sub
            End If
        '
        '    ' Check to see if player is already in a party
        '    If player(n).InParty = NO Then
        If player(index).Party = 0 Then
        i = FindOpenParty
        If i > 0 Then
        player(index).Party = i
        Parties(i).PartyLeader = index
        Parties(i).Player1 = index
        Parties(i).NumParty = 1
        Else
        Call PlayerMsgCombat(index, "Parties are full for some reason...", White)
        Exit Sub
        End If
        End If
        player(n).PartyPlayer = index
        
        player(index).InParty = YES
        
                Call PlayerMsgCombat(index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsgCombat(n, GetPlayerName(index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
        '
        '        player(Index).PartyStarter = YES
        '
        '        If player(Index).nPartyPlayer = 0 Then
        '
        '        player(Index).nPartyPlayer = n
        '        ElseIf player(Index).nPartyPlayer2 = 0 Then
        '
        '        player(Index).nPartyPlayer2 = n
        '        ElseIf player(Index).nPartyPlayer3 = 0 Then
        '
        '        player(Index).nPartyPlayer3 = n
        '        Else
        '        Call PlayerMsg(Index, "Party is full! or to many people pending.", BrightRed)
        '        Exit Sub
        '        End If
        '        player(n).PartyPlayer = Index
        '        player(Index).InParty = YES
        '        player(n).InParty = YES
        '    Else
        '        Call PlayerMsg(Index, "Player is already in a party!", Pink)
        '    End If
            
        
        
        Else
            Call PlayerMsgCombat(index, "Player is not online.", White)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(parse(0)) = "joinparty" Then
        n = player(index).PartyPlayer
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            'If player(Index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
            '    If player(n).nPartyPlayer = Index Or player(n).nPartyPlayer2 = Index Or player(n).nPartyPlayer3 = Index Then
            '        If player(n).PartyNum >= 3 Then
            '        Call PlayerMsg(Index, "PARTY FULL!", BrightRed)
            '         If player(n).nPartyPlayer = Index Then
            '        player(n).nPartyPlayer = 0
            '        End If
            '        If player(n).nPartyPlayer2 = Index Then
            '        player(n).nPartyPlayer = 0
            '        End If
            '        If player(n).nPartyPlayer3 = Index Then
            '        player(n).nPartyPlayer = 0
            '        End If
            '        Exit Sub
            '        End If
            If Parties(player(n).Party).NumParty <= 3 Then
            If player(index).Party = 0 Then
                    Call PlayerMsgCombat(index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                    Call PlayerMsgCombat(n, GetPlayerName(index) & " has joined your party!", Pink)
                    player(index).Party = player(n).Party
                    Parties(player(index).Party).NumParty = Parties(player(index).Party).NumParty + 1
                    If Parties(player(index).Party).Player1 = 0 Then
                    Parties(player(index).Party).Player1 = index
                    ElseIf Parties(player(index).Party).Player2 = 0 Then
                    Parties(player(index).Party).Player2 = index
                    ElseIf Parties(player(index).Party).Player3 = 0 Then
                    Parties(player(index).Party).Player3 = index
                    ElseIf Parties(player(index).Party).Player4 = 0 Then
                    Parties(player(index).Party).Player4 = index
                    End If
                    player(index).InParty = YES
                    player(index).PartyPlayer = 0
                    
             '       If player(n).PartyPlayer = 0 Then
             '       player(n).PartyPlayer = Index
             '       'player(Index).PartyPlayer = n
             '       player(n).PartyNum = player(n).PartyNum + 1
             '       ElseIf player(n).PartyPlayer2 = 0 Then
             '       player(n).PartyPlayer2 = Index
             '       'player(Index).PartyPlayer = n
             '       player(n).PartyNum = player(n).PartyNum + 1
             '       ElseIf player(n).PartyPlayer3 = 0 Then
              '      player(n).PartyPlayer3 = Index
              ''      'player(Index).PartyPlayer = n
               '     player(n).PartyNum = player(n).PartyNum + 1
               ''     Else
               '     Call PlayerMsg(Index, "Party failed.", Pink)
               '      If player(n).nPartyPlayer = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '     If player(n).nPartyPlayer2 = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '     If player(n).nPartyPlayer3 = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '     Exit Sub
               '     End If
               '     player(Index).InParty = YES
               '     player(n).InParty = YES
               '     If player(n).nPartyPlayer = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '     If player(n).nPartyPlayer2 = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '     If player(n).nPartyPlayer3 = Index Then
               '     player(n).nPartyPlayer = 0
               '     End If
               '
                Else
                    Call PlayerMsgCombat(index, "Your already in a party!", Pink)
                End If
            Else
                Call PlayerMsgCombat(index, "Party is Full!", Pink)
            End If
        Else
            Call PlayerMsgCombat(index, "You have not been invited into a party!", Pink)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    If LCase(parse(0)) = "leaveparty" Then
        n = player(index).PartyPlayer
        i = player(index).Party
        
        If n > 0 Or i > 0 Then
            If player(index).InParty = YES Then
                player(index).InParty = NO
                
                Call PlayerMsgCombat(index, "You have left the party.", Pink)
                'Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
               ' If player(Index).PartyStarter = YES Then
                
                
               ' Call PlayerMsg(player(Index).PartyPlayer, "Group has been disbanded!", BrightRed)
               ' Call PlayerMsg(player(Index).PartyPlayer2, "Group has been disbanded!", BrightRed)
               ' Call PlayerMsg(player(Index).PartyPlayer3, "Group has been disbanded!", BrightRed)
               ' player(Index).InParty = 0
               ' player(player(Index).PartyPlayer).PartyPlayer = 0
               ' player(player(Index).PartyPlayer2).PartyPlayer = 0
               ' player(player(Index).PartyPlayer3).PartyPlayer = 0
               ' player(player(Index).PartyPlayer).InParty = 0
               ' player(player(Index).PartyPlayer2).InParty = 0
               ' player(player(Index).PartyPlayer3).InParty = 0
               ' player(Index).PartyPlayer = 0
               ' player(Index).PartyPlayer2 = 0
               ' player(Index).PartyPlayer3 = 0
               ' player(Index).PartyNum = 0
               ' Else
               ' player(Index).InParty = 0
               ' If player(player(Index).PartyPlayer).PartyPlayer = Index Then
               ' player(player(Index).PartyPlayer).PartyPlayer = 0
               ' player(player(Index).PartyPlayer).PartyNum = player(player(Index).PartyPlayer).PartyNum - 1
               ' ElseIf player(player(Index).PartyPlayer).PartyPlayer2 = Index Then
               ' player(player(Index).PartyPlayer).PartyPlayer2 = 0
               ' player(player(Index).PartyPlayer).PartyNum = player(player(Index).PartyPlayer).PartyNum - 1
               ' ElseIf player(player(Index).PartyPlayer).PartyPlayer3 = Index Then
               ' player(player(Index).PartyPlayer).PartyPlayer3 = 0
               '
               ' player(player(Index).PartyPlayer).PartyNum = player(player(Index).PartyPlayer).PartyNum - 1
               ' End If
               ' End If
               '
               '
               If Parties(i).NumParty >= 1 Then
               Parties(i).NumParty = Parties(i).NumParty - 1
               End If
               If Parties(i).Player1 = index Then
               Parties(i).Player1 = 0
               ElseIf Parties(i).Player2 = index Then
               Parties(i).Player2 = 0
                ElseIf Parties(i).Player3 = index Then
               Parties(i).Player3 = 0
                ElseIf Parties(i).Player4 = index Then
               Parties(i).Player4 = 0
               End If
               If Parties(i).PartyLeader = index Then
               If Parties(i).Player1 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player1
               Call PlayerMsgCombat(Parties(i).Player1, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player2 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player2
               Call PlayerMsgCombat(Parties(i).Player2, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player3 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player3
               Call PlayerMsgCombat(Parties(i).Player3, "You are now the leader of the party!", BrightRed)
               ElseIf Parties(i).Player4 > 0 Then
               Parties(i).PartyLeader = Parties(i).Player4
               Call PlayerMsgCombat(Parties(i).Player4, "You are now the leader of the party!", BrightRed)
               End If
               End If
               player(index).Party = 0
               
               
            Else
                Call PlayerMsgCombat(index, "Declined party request.", Pink)
                Call PlayerMsgCombat(n, GetPlayerName(index) & " declined your request.", Pink)
            
                player(index).PartyPlayer = 0
                player(index).PartyStarter = NO
                player(index).InParty = NO
          '       If player(n).nPartyPlayer = Index Then
          '          player(n).nPartyPlayer = 0
          '          End If
           '        If player(n).nPartyPlayer2 = Index Then
           '        player(n).nPartyPlayer = 0
           '        End If
           '        If player(n).nPartyPlayer3 = Index Then
           '         player(n).nPartyPlayer = 0
           '         End If
            End If
        Else
           Call PlayerMsgCombat(index, "You are not in a party!", Pink)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If LCase(parse(0)) = "spells" Then
        Call SendPlayerSpells(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(parse(0)) = "cast" Then
        ' Spell slot
        n = Val(parse(1))
        
        Call CastSpell(index, n)
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Forget Spell ::
    ' ::::::::::::::::::
    If LCase(parse(0)) = "forgetspell" Then
    n = Val(parse(1))
        
    Call SetPlayerSpell(index, n, 0)
    Exit Sub
    End If
    

    ' ::::::::::::::::
    ' :: Bank Packet::
    ' ::::::::::::::::
    If LCase(parse(0)) = "bank" Then
        If MAP(GetPlayerMap(index)).BootY = 1 Then
        Call SendBank(index)
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::::::::
    ' :: Bank Take request packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(parse(0)) = "takebank" Then
    If MAP(GetPlayerMap(index)).BootY <> 1 Then
    Call PlayerMsgCombat(index, "NO BANK HERE!", BrightRed)
    Exit Sub
    End If
        ' Trade num
        n = Val(parse(1))
        i = Val(parse(2))
        If i > GetPlayerBankItemValue(index, n) Then
        i = GetPlayerBankItemValue(index, n)
        End If
        
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_BANK) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Check if inv full
        X = FindOpenInvSlot(index, GetPlayerBankItemNum(index, n))
        If X = 0 Then
            Call PlayerMsgCombat(index, "Withdrawl unsuccessful, Bank full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        'If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            'Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            If Item(GetPlayerBankItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveItem(index, GetPlayerBankItemNum(index, n), i)
            Call TakeBankItem(index, GetPlayerBankItemNum(index, n), i)
            Else
            Call GiveItem(index, GetPlayerBankItemNum(index, n), GetPlayerBankItemValue(index, n))
            Call TakeBankItem(index, GetPlayerBankItemNum(index, n), GetPlayerBankItemValue(index, n))
            Call PlayerMsgCombat(index, "Withdrawl was successful!", Yellow)
            End If
        'Else
        '    Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        'End If
        Exit Sub
    End If
    'Deposit into bank
    If LCase(parse(0)) = "givebank" Then
    If MAP(GetPlayerMap(index)).BootY <> 1 Then
    Call PlayerMsgCombat(index, "NO BANK HERE!", BrightRed)
    Exit Sub
    End If
        ' Trade num
        n = Val(parse(1))
        i = Val(parse(2))
        If i > GetPlayerInvItemValue(index, n) Then
        i = GetPlayerInvItemValue(index, n)
        End If
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_BANK) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Check if inv full
        X = FindOpenBankSlot(index, GetPlayerInvItemNum(index, n))
        If X = 0 Then
            Call PlayerMsgCombat(index, "Deposit unsuccessful, Bank full.", BrightRed)
            Exit Sub
        End If
        InvNum = n
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Weapon before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Accessory before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Armor before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Shield before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        
        ' Check if they have the item
        'If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            'Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveBankItem(index, GetPlayerInvItemNum(index, n), i)
            Call TakeItem(index, GetPlayerInvItemNum(index, n), i)
            Else
            
            'Call SetPlayerSTR(Index, (GetPlayerSTR(Index) - Item(GetPlayerInvItemNum(Index, n)).STRmod))
            'Call SetPlayerDEF(Index, (GetPlayerDEF(Index) - Item(GetPlayerInvItemNum(Index, n)).DEFmod))
            'Call SetPlayerMAGI(Index, (GetPlayerMAGI(Index) - Item(GetPlayerInvItemNum(Index, n)).MAGImod))
            Call GiveBankItem(index, GetPlayerInvItemNum(index, n), GetPlayerInvItemValue(index, n))
            Call TakeItem(index, GetPlayerInvItemNum(index, n), GetPlayerInvItemValue(index, n))
            
            
            Call PlayerMsgCombat(index, "Deposit was successful!", Yellow)
            End If
        'Else
        '    Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        'End If
        Exit Sub
    End If

      If LCase(parse(0)) = "giveitem" Then
      Dim IT As Long
    IT = player(index).Target
    If player(index).TargetType = TARGET_TYPE_NPC Then
    Call PlayerMsgCombat(index, "You cant give items to npc's....", BrightRed)
    Exit Sub
    End If
    
    
    
     ' Trade num
        n = Val(parse(1))
        i = Val(parse(2))
        If i > GetPlayerInvItemValue(index, n) Then
        i = GetPlayerInvItemValue(index, n)
        End If
        
        
    X = FindOpenInvSlot(IT, GetPlayerInvItemNum(index, n))
    If X = 0 Then
            Call PlayerMsgCombat(index, "Trade unsuccessful, Inventory full.", BrightRed)
            Exit Sub
        End If
    If Item(GetPlayerInvItemNum(index, n)).NoDrop > 0 Then
    Call PlayerMsgCombat(index, "This item is a No Drop item. You can not trade it with players", BrightRed)
    Exit Sub
    End If
           InvNum = n
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Weapon before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Accessory before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Armor before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        tEmp2 = GetPlayerInvItemNum(index, InvNum)
        tEmp3 = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
        If tEmp2 = tEmp3 Then
        Call PlayerMsgCombat(index, "Please unequip your Shield before selling an item of this type.", BrightRed)
        Exit Sub
        End If
        
        ' Check if they have the item
        'If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            'Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then

            Call GiveItem(IT, GetPlayerInvItemNum(index, n), i)
            Call TakeItem(index, GetPlayerInvItemNum(index, n), i)
            Call PlayerMsgCombat(index, "Trade was successful!", Yellow)
            Call PlayerMsgCombat(IT, GetPlayerName(index) & " has sent you " & i & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name), Yellow)
            Else
            
            'Call SetPlayerSTR(Index, (GetPlayerSTR(Index) - Item(GetPlayerInvItemNum(Index, n)).STRmod))
            'Call SetPlayerDEF(Index, (GetPlayerDEF(Index) - Item(GetPlayerInvItemNum(Index, n)).DEFmod))
            'Call SetPlayerMAGI(Index, (GetPlayerMAGI(Index) - Item(GetPlayerInvItemNum(Index, n)).MAGImod))
            Call GiveItem(IT, GetPlayerInvItemNum(index, n), GetPlayerInvItemValue(index, n))
            Call TakeItem(index, GetPlayerInvItemNum(index, n), GetPlayerInvItemValue(index, n))
            
            
            Call PlayerMsgCombat(index, "Trade was successful!", Yellow)
            Call PlayerMsgCombat(IT, GetPlayerName(index) & " has sent you a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name), Yellow)
            
            End If
        'Else
        '    Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        'End If
        Exit Sub
    End If
    
      
    
    
     
    
    
    
    
    
    
    
    
    



    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    If LCase(parse(0)) = "requestlocation" Then
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsgCombat(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), Pink)
        Exit Sub
    End If
End Sub

Sub CloseSocket(ByVal index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        
        frmServer.Socket(index).Close
            
        Call UpdateCaption
        Call ClearPlayer(index)
    End If
End Sub

Sub SendWhosOnline(ByVal index As Long)
Dim s As String
Dim n As Long, i As Long

    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsgCombat(index, s, WhoColor)
End Sub
Sub SendGuildWhosOnline(ByVal index As Long)
Dim s As String
Dim n As Long, i As Long

    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerGuild(i) = GetPlayerGuild(index) Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players from your guild online: " & s & "."
    End If
        
    Call PlayerMsgCombat(index, s, WhoColor)
End Sub

Sub SendChars(ByVal index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(player(index).Char(i).Name) & SEP_CHAR & Trim(Class(player(index).Char(i).Class).Name) & SEP_CHAR & Trim(Race(player(index).Char(i).Race).Name) & SEP_CHAR & player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerSprite2(i) & SEP_CHAR & GetPlayerSprite3(i) & SEP_CHAR & GetPlayerSprite4(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & GetPlayerGlobal(i) & SEP_CHAR & GetPet(i) & SEP_CHAR & GetPlayerGuild(i) & SEP_CHAR & GetPlayerAnonymous(i) & SEP_CHAR & GetPlayerFaction(i) & SEP_CHAR & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i

    ' Send index's player data to everyone on the map including himself

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerSprite2(index) & SEP_CHAR & GetPlayerSprite3(index) & SEP_CHAR & GetPlayerSprite4(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGlobal(index) & SEP_CHAR & GetPet(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerAnonymous(index) & SEP_CHAR & GetPlayerFaction(index) & SEP_CHAR & END_CHAR
      
    
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerSprite2(index) & SEP_CHAR & GetPlayerSprite3(index) & SEP_CHAR & GetPlayerSprite4(index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGlobal(index) & SEP_CHAR & GetPet(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerAnonymous(index) & SEP_CHAR & GetPlayerFaction(index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerSprite2(index) & SEP_CHAR & GetPlayerSprite3(index) & SEP_CHAR & GetPlayerSprite4(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGlobal(index) & SEP_CHAR & GetPet(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerAnonymous(index) & SEP_CHAR & GetPlayerFaction(index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub
Sub SendPlayerStats(ByVal index As Long)
Dim Packet As String
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(MAP(MapNum).Name) & SEP_CHAR & MAP(MapNum).Revision & SEP_CHAR & MAP(MapNum).Moral & SEP_CHAR & MAP(MapNum).Up & SEP_CHAR & MAP(MapNum).Down & SEP_CHAR & MAP(MapNum).Left & SEP_CHAR & MAP(MapNum).Right & SEP_CHAR & MAP(MapNum).Music & SEP_CHAR & MAP(MapNum).BootMap & SEP_CHAR & MAP(MapNum).BootX & SEP_CHAR & MAP(MapNum).BootY & SEP_CHAR & MAP(MapNum).Shop & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With MAP(MapNum).Tile(X, y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
            End With
        Next X
    Next y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & MAP(MapNum).Npc(X) & SEP_CHAR
    Next X
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendItems(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendAdminCmds(ByVal index As Long)
Dim Packet As String
    Packet = "ADMINCMDS" & SEP_CHAR & player(index).Char(player(index).charnum).AdminCmds & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub


Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub
Sub SendBank(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK
        Packet = Packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub
Sub SendBankUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERBANKUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWelcome(ByVal index As Long)
Dim MOTD As String
Dim f As Long

    ' Send them welcome
    Call PlayerMsgCombat(index, "Welcome to " & GAME_NAME & "! Version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, BrightBlue)
    Call PlayerMsgCombat(index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    ' Send them MOTD
    MOTD = GetVar(App.Path & "\motd.ini", "MOTD", "Msg")
    If Trim(MOTD) <> "" Then
        Call PlayerMsgCombat(index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPeed & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendRaces(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "RACESDATA" & SEP_CHAR & Max_Races & SEP_CHAR
    For i = 0 To Max_Races
        Packet = Packet & GetRaceName(i) & SEP_CHAR & GetRaceMaxHP(i) & SEP_CHAR & GetRaceMaxMP(i) & SEP_CHAR & GetRaceMaxSP(i) & SEP_CHAR & Race(i).STR & SEP_CHAR & Race(i).DEF & SEP_CHAR & Race(i).SPeed & SEP_CHAR & Race(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub


Sub SendNewCharClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPeed & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharRaces(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARRACES" & SEP_CHAR & Max_Races & SEP_CHAR
    For i = 0 To Max_Races
        Packet = Packet & GetRaceName(i) & SEP_CHAR & GetRaceMaxHP(i) & SEP_CHAR & GetRaceMaxMP(i) & SEP_CHAR & GetRaceMaxSP(i) & SEP_CHAR & Race(i).STR & SEP_CHAR & Race(i).DEF & SEP_CHAR & Race(i).SPeed & SEP_CHAR & Race(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendDataToOoc(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerOocSwitch(i) >= 1 Then
            
            Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub
Sub SendDataTotc(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayertcSwitch(i) >= 1 Then
            
            Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendLeftGame(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR
    Call SendDataToAllBut(index, Packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub
Sub GlobalMsgOoc(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSGGLOBAL" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToOoc(Packet)
End Sub
Sub GlobalMsg2(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSGGLOBAL" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToOoc(Packet)
End Sub
Sub GlobalMsgtc(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSGGLOBAL" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTotc(Packet)
End Sub
Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).Class & SEP_CHAR & Item(ItemNum).spdmod & SEP_CHAR & Item(ItemNum).STRmod & SEP_CHAR & Item(ItemNum).DEFmod & SEP_CHAR & Item(ItemNum).MAGImod & SEP_CHAR & Item(ItemNum).SPRITE & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & Item(ItemNum).NoDrop & SEP_CHAR & END_CHAR
    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Class & SEP_CHAR & Item(ItemNum).STRmod & SEP_CHAR & Item(ItemNum).DEFmod & SEP_CHAR & Item(ItemNum).MAGImod & SEP_CHAR & Item(ItemNum).SPRITE & SEP_CHAR & Item(ItemNum).SellValue & SEP_CHAR & Item(ItemNum).NoDrop & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).SPRITE & SEP_CHAR & Npc(NpcNum).SPRITE2 & SEP_CHAR & Npc(NpcNum).SPRITE3 & SEP_CHAR & Npc(NpcNum).SPRITE4 & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).SPRITE & SEP_CHAR & Npc(NpcNum).SPRITE2 & SEP_CHAR & Npc(NpcNum).SPRITE3 & SEP_CHAR & Npc(NpcNum).SPRITE4 & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).SPRITE & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPeed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).DropChance2 & SEP_CHAR & Npc(NpcNum).DropItem2 & SEP_CHAR & Npc(NpcNum).DropItemValue2 & SEP_CHAR & Npc(NpcNum).SPRITE2 & SEP_CHAR & Npc(NpcNum).SPRITE3 & SEP_CHAR & Npc(NpcNum).SPRITE4 & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendShops(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).OneSale & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub
Sub SendItemInvDetails(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "itemdetails" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(index, ItemNum)).Name) & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Pic & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Type & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Data1 & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Data2 & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Data3 & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).Class & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).STRmod & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).DEFmod & SEP_CHAR & Item(GetPlayerInvItemNum(index, ItemNum)).MAGImod & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub
Sub SendItemShopDetails(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String
Dim i As Long
i = MAP(GetPlayerMap(index)).Shop
    Packet = "shopitemdetails" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item((Shop(i).TradeItem(ItemNum).GetItem)).Name) & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Pic & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Type & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Data1 & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Data2 & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Data3 & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).Class & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).STRmod & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).DEFmod & SEP_CHAR & Item(Shop(i).TradeItem(ItemNum).GetItem).MAGImod & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSpells(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
End Sub
Sub SendDataToGuild(ByVal index As Long, Data As String)
Dim i As Long
For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerGuild(index) = GetPlayerGuild(i) Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub
Sub SendDataToParty(ByVal index As Long, Data As String)
Dim i As Long
Dim n As Long
If player(index).Party > 0 Then
n = player(index).Party
If Parties(n).Player1 > 0 Then
Call SendDataTo(Parties(n).Player1, Data)
End If
If Parties(n).Player2 > 0 Then
Call SendDataTo(Parties(n).Player2, Data)
End If

If Parties(n).Player3 > 0 Then
Call SendDataTo(Parties(n).Player3, Data)
End If

If Parties(n).Player4 > 0 Then
Call SendDataTo(Parties(n).Player4, Data)
End If
End If


End Sub
Sub GlobalMsgGuild(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToGuild(index, Packet)
End Sub
Sub GlobalMsgParty(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToParty(index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(spell(SpellNum).Name) & SEP_CHAR & spell(SpellNum).Sfx & SEP_CHAR & spell(SpellNum).LevelReq & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(spell(SpellNum).Name) & SEP_CHAR & spell(SpellNum).Sfx & SEP_CHAR & spell(SpellNum).LevelReq & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(spell(SpellNum).Name) & SEP_CHAR & spell(SpellNum).ClassReq & SEP_CHAR & spell(SpellNum).LevelReq & SEP_CHAR & spell(SpellNum).Type & SEP_CHAR & spell(SpellNum).Data1 & SEP_CHAR & spell(SpellNum).Data2 & SEP_CHAR & spell(SpellNum).Data3 & SEP_CHAR & spell(SpellNum).MPused & SEP_CHAR & spell(SpellNum).Sfx & SEP_CHAR & spell(SpellNum).Gfx & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, X As Long, y As Long

    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).OneSale & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
        
        ' Item #
        X = Shop(ShopNum).TradeItem(i).GetItem
        
        'If Item(X).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
           ' y = spell(Item(X).Data1).ClassReq
            
          '  If y = 0 Then
          '      Call PlayerMsg(Index, Trim(Item(X).Name) & " can be used by all classes.", Yellow)
           ' Else
          '      Call PlayerMsg(Index, Trim(Item(X).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
          '  End If
       ' End If
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendPlayerSpells(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendTimeTo(ByVal index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
End Sub

