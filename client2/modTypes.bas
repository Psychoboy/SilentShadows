Attribute VB_Name = "modTypes"
Option Explicit
Public emotetemp As Byte
' Winsock globals
Public Const GAME_PORT = 7555
Public Const GAME_IP = "67.19.111.154"
'Public Const GAME_IP = "mercury.rilphiasaga.com"
'Public Const GAME_IP = "192.168.0.2"
Public TempStr As String
Public tEmp As Integer
Public GLBLCHT As Integer
Public IHaveMusic As Integer
Public IHaveSFX As Integer

Public SHeight As Single
Public SWidth As Single

Public SELectINV

Public TEMPINVSLOT As Integer
Public TEMPITEMNUM As Integer

Public QUICKSPELL1 As Integer
Public QUICKSPELL2 As Integer
Public QUICKSPELL3 As Integer


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long




'Call mciSendString("status song mode", strReturn, Len(strReturn), 0)
'if instr(strReturn, "playing") > 0 then Playing = True: Else: Playing = False

' General constants
'"status sequencer mode"
Public strReturn2 As String * 255
Public Playing As Boolean
Public MyMusiC As Integer

Public Const GAME_NAME = "Silent Shadows"
Public Const MAX_PLAYERS = 200
Public Const MAX_ITEMS = 1000
Public Const MAX_NPCS = 500
Public Const MAX_INV = 50
Public Const MAX_BANK = 50
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_MAP_NPCS = 15
Public Const MAX_SHOPS = 255
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_SPELLS = 255
Public Const MAX_TRADES = 15

Public Const NO = 0
Public Const YES = 1

Public BuDDy(10) As String

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
Public Const MAX_MAPS = 10000
Public Const MAX_MAPX = 15
Public Const MAX_MAPY = 11
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_ARENA = 2
Public Const MAP_MORAL_HOUSE = 3

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_SIGN = 7
Public Const TILE_TYPE_FULLHEAL = 8
Public Const TILE_TYPE_DEATH = 9
Public Const TILE_TYPE_SFX = 10
Public Const TILE_TYPE_QUEST = 11
Public Const TILE_TYPE_ARENA = 12
Public Const TILE_TYPE_FISHING = 13
Public Const TILE_TYPE_MINING = 14

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13
Public Const ITEM_TYPE_SCHANGE = 14
Public Const ITEM_TYPE_BOOK = 15
Public Const ITEM_TYPE_RANDITEM = 16
Public Const ITEM_TYPE_MASKEDIT = 17
Public Const ITEM_TYPE_PETEGG = 18
Public Const ITEM_TYPE_BARBER = 19

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4
Public Const ADMIN_LORD = 5

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_CRAFT = 7

Type PlayerInvRec
    num As Integer
    value As Long
    Dur As Long
End Type

Type PlayerBankRec
    num As Integer
    value As Long
    Dur As Long
End Type

Type PlayerRec
    ' General
    name As String * NAME_LENGTH
    Class As Byte
    Race As Byte
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    spellsprite As Integer
    spellframe As Integer
    emotesprite As Integer
    emoteframe As Integer
    PET As Integer
    Level As Long
    Exp As Long
    Access As Byte
    AdminCmds As String
    PK As Byte
    guild As String * NAME_LENGTH
    GuildRank As Byte
    Faction As Byte
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    
    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Fishing As Long
    Mining As Long
    MAGI As Long
    POINTS As Long
    ANONYMOUS As Byte
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Bank(1 To MAX_BANK) As PlayerBankRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
       
    ' Position
    Map As Integer
    X As Byte
    y As Byte
    Dir As Byte
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
End Type
    
Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    M2Anim As Integer
    Fringe As Integer
    FAnim As Integer
    Fringe2 As Integer
    F2Anim As Integer
    Type As Byte
    Data1 As String
    Data2 As String
    Data3 As String
End Type

Type MapRec
    name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Integer
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Integer
End Type

Type RaceRec
    name As String * NAME_LENGTH
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    
    STR As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ClassRec
    name As String * NAME_LENGTH
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    
    STR As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    name As String * NAME_LENGTH
    Class As Integer
    Pic As Integer
    Type As Byte
    Data1 As String
    Data2 As String
    Data3 As Integer
    Data4 As Integer
    strmod As Integer
    defmod As Integer
    magimod As Integer
    SPRITE As Integer
    spdmod As Integer
    SellValue As Long
    NoDrop As Byte
    Sfx As String
End Type

Type MapItemRec
    num As Integer
    value As Long
    Dur As Integer
    
    
    X As Byte
    y As Byte
End Type

Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 300
    
    SPRITE As Long
    SPRITE2 As Long
    SPRITE3 As Long
    SPRITE4 As Long
    spellsprite As Integer
    spellframe As Integer
    SpawnSecs As Long
    Behavior As Byte
    range As Byte
    
    DropChance As Integer
    DropItem As Integer
    DropItemValue As Integer
    
    DropChance2 As Integer
    DropItem2 As Integer
    DropItemValue2 As Integer
    
    STR  As Integer
    DEF As Integer
    Speed As Integer
    MAGI As Integer
End Type

Type MapNpcRec
    num As Integer
    
    Target As Byte
    
    HP As Long
    MP As Long
    SP As Long
        
    Map As Integer
    X As Byte
    y As Byte
    Dir As Byte
    spellsprite As Integer
    spellframe As Integer
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    OneSale As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Type SpellRec
    name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    MPused As Integer
    range As Integer
    Usedef As Boolean
    Sfx As String
    GFX As Integer
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public Max_Races As Byte

Public AdminCmds As String
Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Race() As RaceRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec

Sub ClearTempTile()
Dim X As Long, y As Long

    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, y).DoorOpen = NO
        Next X
    Next y
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim i As Long
Dim n As Long

    Player(index).name = ""
    Player(index).Class = 0
    Player(index).Race = 0
    Player(index).Level = 0
    Player(index).SPRITE = 0
    Player(index).Exp = 0
    Player(index).Access = 0
    Player(index).PK = NO
    Player(index).guild = ""
    Player(index).Faction = 0
        
    Player(index).HP = 0
    Player(index).MP = 0
    Player(index).SP = 0
        
    Player(index).STR = 0
    Player(index).DEF = 0
    Player(index).Speed = 0
    Player(index).MAGI = 0
        
    For n = 1 To MAX_INV
        Player(index).Inv(n).num = 0
        Player(index).Inv(n).value = 0
        Player(index).Inv(n).Dur = 0
    Next n
        
    Player(index).ArmorSlot = 0
    Player(index).WeaponSlot = 0
    Player(index).HelmetSlot = 0
    Player(index).ShieldSlot = 0
        
    Player(index).Map = 0
    Player(index).X = 0
    Player(index).y = 0
    Player(index).Dir = 0
    
    ' Client use only
    Player(index).MaxHP = 0
    Player(index).MaxMP = 0
    Player(index).MaxSP = 0
    Player(index).XOffset = 0
    Player(index).YOffset = 0
    Player(index).Moving = 0
    Player(index).Attacking = 0
    Player(index).AttackTimer = 0
    Player(index).MapGetTimer = 0
    Player(index).CastedSpell = NO
End Sub

Sub ClearItem(ByVal index As Long)
    Item(index).name = ""
    
    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long)
    MapItem(index).num = 0
    MapItem(index).value = 0
    MapItem(index).Dur = 0
    MapItem(index).X = 0
    MapItem(index).y = 0
End Sub

Sub ClearMap()
Dim i As Long
Dim X As Long
Dim y As Long

    Map.name = ""
    Map.Revision = 0
    Map.Moral = 0
    Map.Up = 0
    Map.Down = 0
    Map.Left = 0
    Map.Right = 0
        
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map.Tile(X, y).Ground = 0
            Map.Tile(X, y).Mask = 0
            Map.Tile(X, y).Anim = 0
            Map.Tile(X, y).Mask2 = 0
            Map.Tile(X, y).M2Anim = 0
            Map.Tile(X, y).Fringe = 0
            Map.Tile(X, y).FAnim = 0
            Map.Tile(X, y).Fringe2 = 0
            Map.Tile(X, y).F2Anim = 0
            Map.Tile(X, y).Type = 0
            Map.Tile(X, y).Data1 = 0
            Map.Tile(X, y).Data2 = 0
            Map.Tile(X, y).Data3 = 0
        Next X
    Next y
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub
Sub MusicTest()
If MIDIPlaying = True Then
Playing = True
Else
Playing = False
End If
If Playing = False Then
Call PlayMidi(MyMusiC & ".ogg")
End If
End Sub
Sub ClearMapNpc(ByVal index As Long)
    MapNpc(index).num = 0
    MapNpc(index).Target = 0
    MapNpc(index).HP = 0
    MapNpc(index).MP = 0
    MapNpc(index).SP = 0
    MapNpc(index).Map = 0
    MapNpc(index).X = 0
    MapNpc(index).y = 0
    MapNpc(index).Dir = 0
    
    ' Client use only
    MapNpc(index).XOffset = 0
    MapNpc(index).YOffset = 0
    MapNpc(index).Moving = 0
    MapNpc(index).Attacking = 0
    MapNpc(index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim(Player(index).name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    Player(index).name = name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerRace(ByVal index As Long) As Long
    GetPlayerRace = Player(index).Race
End Function

Sub SetPlayerRace(ByVal index As Long, ByVal RaceNum As Long)
    Player(index).Race = RaceNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).SPRITE
End Function
Function GetPlayerSprite2(ByVal index As Long) As Long
    GetPlayerSprite2 = Player(index).SPRITE2
End Function
Function GetPlayerSprite3(ByVal index As Long) As Long
    GetPlayerSprite3 = Player(index).SPRITE3
End Function
Function GetPlayerSprite4(ByVal index As Long) As Long
    GetPlayerSprite4 = Player(index).SPRITE4
End Function
Function GetPet(ByVal index As Long) As Long
    GetPet = Player(index).PET
End Function
Function Getspell(ByVal index As Long) As Long
    Getspell = Player(index).spellsprite
End Function
Function Getframe(ByVal index As Long) As Long
    Getframe = Player(index).spellframe
End Function
Function GetEmoteSprite(ByVal index As Long) As Long
    GetEmoteSprite = Player(index).emotesprite
End Function
Function GetEmoteframe(ByVal index As Long) As Long
    GetEmoteframe = Player(index).emoteframe
End Function
Function Getspell2(ByVal index As Long) As Long
    Getspell2 = MapNpc(index).spellsprite
End Function
Function Getframe2(ByVal index As Long) As Long
    Getframe2 = MapNpc(index).spellframe
End Function
Sub SetSpell(ByVal index As Long, ByVal SPRITE As Long)
Player(index).spellsprite = SPRITE
End Sub
Sub SetFrame(ByVal index As Long, ByVal SPRITE As Long)
Player(index).spellframe = SPRITE
End Sub
Sub SetEmote(ByVal index As Long, ByVal SPRITE As Long)
Player(index).emotesprite = SPRITE
End Sub
Sub SetEmoteFrame(ByVal index As Long, ByVal SPRITE As Long)
Player(index).emoteframe = SPRITE
End Sub
Sub SetSpell2(ByVal index As Long, ByVal SPRITE As Long)
If index <= 0 Then
        Exit Sub
    End If
MapNpc(index).spellsprite = SPRITE
End Sub
Sub SetFrame2(ByVal index As Long, ByVal SPRITE As Long)
MapNpc(index).spellframe = SPRITE
End Sub
Sub SetPlayerSprite(ByVal index As Long, ByVal SPRITE As Long)
    Player(index).SPRITE = SPRITE
End Sub
Sub SetPlayerSprite2(ByVal index As Long, ByVal SPRITE As Long)
    Player(index).SPRITE2 = SPRITE
End Sub
Sub SetPlayerSprite3(ByVal index As Long, ByVal SPRITE As Long)
    Player(index).SPRITE3 = SPRITE
End Sub
Sub SetPlayerSprite4(ByVal index As Long, ByVal SPRITE As Long)
    Player(index).SPRITE4 = SPRITE
End Sub
Sub SetPet(ByVal index As Long, ByVal SPRITE As Long)
    Player(index).PET = SPRITE
End Sub
Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Level = Level
End Sub

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).HP = GetPlayerMaxHP(index)
    End If
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).MP = GetPlayerMaxMP(index)
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).SP = GetPlayerMaxSP(index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
    GetPlayerMaxHP = Player(index).MaxHP
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
    GetPlayerMaxMP = Player(index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
    GetPlayerMaxSP = Player(index).MaxSP
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
    GetPlayerSTR = Player(index).STR
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = Player(index).DEF
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = Player(index).Speed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)
    Player(index).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    Player(index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).POINTS
End Function
Function GetPlayerAnonymous(ByVal index As Long) As Byte
    GetPlayerAnonymous = Player(index).ANONYMOUS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).POINTS = POINTS
End Sub
Sub SetPlayerAnonymous(ByVal index As Long, ByVal POINTS As Byte)
    Player(index).ANONYMOUS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    Player(index).Map = MapNum
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
    Player(index).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Inv(InvSlot).num
End Function
Function GetPlayerBankItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemNum = Player(index).Bank(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Inv(InvSlot).num = ItemNum
End Sub
Sub SetPlayerbankItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Bank(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Inv(InvSlot).value
End Function
Function GetPlayerbankItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerbankItemValue = Player(index).Bank(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(InvSlot).value = ItemValue
End Sub
Sub SetPlayerBankItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Bank(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Inv(InvSlot).Dur
End Function
Function GetPlayerbankItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerbankItemDur = Player(index).Bank(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Inv(InvSlot).Dur = ItemDur
End Sub
Sub SetPlayerbankItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Bank(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).ArmorSlot = InvNum
End Sub
Sub SetPlayerGuild(ByVal index As Long, ByVal guild As String)
    Player(index).guild = guild
End Sub
Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim(Player(index).guild)
End Function
Sub SetPlayerFaction(ByVal index As Long, ByVal guild As Byte)
    Player(index).Faction = guild
End Sub
Function GetPlayerFaction(ByVal index As Long) As Byte
    GetPlayerFaction = Player(index).Faction
End Function
Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).ShieldSlot = InvNum
End Sub
Sub Screenshot()
Dim Cont As Boolean, FilePath As String, Path As String, Increment As Integer
If LCase(Dir(App.Path & "\screenshots", vbDirectory)) <> "screenshots" Then
        Call MkDir(App.Path & "\Screenshots")
    End If
    Path = App.Path & "\Screenshots\"
'Auto Increment
Cont = True
Increment = 1
Do
FilePath = Path & "Screen" & CStr(Increment) & ".bmp"
'There's a better way to determine if a file exists, this is VERY sloppy, in my opinion.
If Dir(FilePath) <> "" Then
    Increment = Increment + 1
Else
    Cont = False
End If
Loop Until Not Cont
SavePicture frmMirage.picScreen.Picture, Path & "Screen" & Increment & ".bmp"

End Sub

Sub CastSpell1()
    If Player(MyIndex).Spell(QUICKSPELL1 + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & QUICKSPELL1 + 1 & SEP_CHAR & END_CHAR)
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

Sub CastSpell2()
    If Player(MyIndex).Spell(QUICKSPELL2 + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & QUICKSPELL2 + 1 & SEP_CHAR & END_CHAR)
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

Sub CastSpell3()
    If Player(MyIndex).Spell(QUICKSPELL3 + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & QUICKSPELL3 + 1 & SEP_CHAR & END_CHAR)
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

Sub DRAWPLAYERUP(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
rec.top = GetPet(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset + 28
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
        Player(index).XOffset = Player(index).XOffset
         Player(index).YOffset = Player(index).YOffset
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
rec.top = GetPlayerSprite3(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    

rec.top = GetPlayerSprite4(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = GetPlayerSprite(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    rec.top = GetPlayerSprite2(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset - 1
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 26
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + 12
        rec.Bottom = rec.Bottom - 12
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
       ' If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

 

End Sub

Sub DRAWPLAYERdown(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
rec.top = GetPet(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 36
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
        Player(index).XOffset = Player(index).XOffset
         Player(index).YOffset = Player(index).YOffset
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    


rec.top = GetPlayerSprite(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
        Player(index).XOffset = Player(index).XOffset
         Player(index).YOffset = Player(index).YOffset
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = GetPlayerSprite2(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 24
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
         
        rec.top = rec.top + 128
         'rec.top = rec.top + (16 - GetPlayerY(index) * PIC_Y + Player(index).YOffset)
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
     '   If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
      '  If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    rec.top = GetPlayerSprite3(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    rec.top = GetPlayerSprite4(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub
Sub DRAWPLAYERleft(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
rec.top = GetPet(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset + 32
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
        Player(index).XOffset = Player(index).XOffset
         Player(index).YOffset = Player(index).YOffset
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
rec.top = GetPlayerSprite4(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = GetPlayerSprite3(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    rec.top = GetPlayerSprite(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    rec.top = GetPlayerSprite3(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

rec.top = GetPlayerSprite2(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset + 2
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 25
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
         rec.top = rec.top + 12
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    

End Sub

Sub DRAWPLAYERright(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
rec.top = GetPet(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset - 32
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
        Player(index).XOffset = Player(index).XOffset
         Player(index).YOffset = Player(index).YOffset
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
     rec.top = GetPlayerSprite3(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
    rec.top = GetPlayerSprite(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   rec.top = GetPlayerSprite2(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 25
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        
         rec.top = rec.top + 12
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
       ' If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = GetPlayerSprite4(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
        If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Sub DRAWNPCright(ByVal MapNpcNum As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    
  rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE3 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
      
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE2 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 25
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE4 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Sub DRAWNPCleft(ByVal MapNpcNum As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE4 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE3 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE2 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 24
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Sub DRAWNPCup(ByVal MapNpcNum As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    
   rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE3 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
 
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE4 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE2 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 26
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
End Sub

Sub DRAWNPCdown(ByVal MapNpcNum As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE2 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 24
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE3 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    rec.top = Npc(MapNpc(MapNpcNum).num).SPRITE4 * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Public Sub ClearISel()

End Sub

Public Sub SaveBuddyX()
Dim filename As String
filename = App.Path & "\Buddy.ini"
Call PutVar(filename, "Buddy1", "Name", BuDDy(1))
Call PutVar(filename, "Buddy2", "Name", BuDDy(2))
Call PutVar(filename, "Buddy3", "Name", BuDDy(3))
Call PutVar(filename, "Buddy4", "Name", BuDDy(4))
Call PutVar(filename, "Buddy5", "Name", BuDDy(5))
Call PutVar(filename, "Buddy6", "Name", BuDDy(6))
Call PutVar(filename, "Buddy7", "Name", BuDDy(7))
Call PutVar(filename, "Buddy8", "Name", BuDDy(8))
Call PutVar(filename, "Buddy9", "Name", BuDDy(9))
Call PutVar(filename, "Buddy10", "Name", BuDDy(10))
End Sub
Public Sub LoadBuddy()
Dim filename As String
filename = App.Path & "\Buddy.ini"
BuDDy(1) = GetVar(filename, "Buddy1", "Name")
BuDDy(2) = GetVar(filename, "Buddy2", "Name")
BuDDy(3) = GetVar(filename, "Buddy3", "Name")
BuDDy(4) = GetVar(filename, "Buddy4", "Name")
BuDDy(5) = GetVar(filename, "Buddy5", "Name")
BuDDy(6) = GetVar(filename, "Buddy6", "Name")
BuDDy(7) = GetVar(filename, "Buddy7", "Name")
BuDDy(8) = GetVar(filename, "Buddy8", "Name")
BuDDy(9) = GetVar(filename, "Buddy9", "Name")
BuDDy(10) = GetVar(filename, "Buddy10", "Name")
End Sub

Sub DRAWspellgfxPLAYER(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    
    If Getspell(index) < 1 Then Exit Sub
    rec.top = Getspell(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = Getframe(index) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call SetFrame(index, Getframe(index) + 1)
    
    If Getframe(index) = 13 Then
    Call SetFrame(index, 0)
    Call SetSpell(index, 0)
    Exit Sub
    End If
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
     '   If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
      '  If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpellSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub
Sub DRAWemotegfxPLAYER(ByVal index As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
    
    If GetEmoteSprite(index) < 1 Then Exit Sub
    rec.top = GetEmoteSprite(index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = GetEmoteframe(index) * PIC_X
    rec.Right = rec.Left + PIC_X
    If emotetemp > 5 Then
    Call SetEmoteFrame(index, GetEmoteframe(index) + 1)
    emotetemp = 0
    End If
    emotetemp = emotetemp + 1
    
    If GetEmoteframe(index) = 13 Then
    Call SetEmoteFrame(index, 0)
    Call SetEmote(index, 0)
    Exit Sub
    End If
    
    X = GetPlayerX(index) * PIC_X + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + Player(index).YOffset + 20
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    'If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9997 Then rec.Bottom = rec.Bottom - 8
     '   If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9998 Then rec.Bottom = rec.Bottom - 11
      '  If Val(Map.Tile(GetPlayerX(index), GetPlayerY(index)).Data3) = 9999 Then rec.Bottom = rec.Bottom - 16
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_EmoteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub
Sub DRAWspellgfxNPC(ByVal MapNpcNum As Integer, ByVal Anim As Integer, ByVal X As Integer, ByVal y As Integer)
   
    rec.top = Getspell2(MapNpcNum) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = Getframe2(MapNpcNum) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    If Getspell2(MapNpcNum) = 0 Then Exit Sub
    If Getframe2(MapNpcNum) = 13 Then
    Call SetFrame2(MapNpcNum, 0)
    Call SetSpell2(MapNpcNum, 0)
    Exit Sub
    End If
    Call SetFrame2(MapNpcNum, Getframe2(MapNpcNum) + 1)
    
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X, y, DD_SpellSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
End Sub
