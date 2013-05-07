Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_SpellSurf As DirectDrawSurface7
Public DD_EmoteSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Spell As DDSURFACEDESC2
Public DDSD_Emote As DDSURFACEDESC2

Public DDSD_BackBuffer As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()
SHeight = Screen.Height
SWidth = Screen.Width
If frmMainMenu.fullscreen.value = 1 Then '

frmMirage.WindowState = 2
frmMirage.BorderStyle = 0
'
frmMirage.WindowState = 2
frmMirage.BorderStyle = 0
End If
    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmMirage.Show
    frmMirage.Timer1.Interval = 1000
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
    
    
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY

    ' Check for files existing
    'If FileExist("sprites.bmp") = False Or FileExist("tiles.bmp") = False Or FileExist("items.bmp") = False Then
       ' Call MsgBox("You dont have the graphics files in the same directory as this executable!", vbOKOnly, GAME_NAME)
        'Call GameDestroy
    'End If
    
    ' Set the key for masks
    key.low = 0
    key.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    'creates gfx at runtime
    'If LCase(Dir(App.Path & "\temp", vbDirectory)) <> "temp" Then
        'Call MkDir(App.Path & "\temp")
       
        
   ' End If
    'SavePicture GFX.items.Picture, App.Path & "\temp\items.bmp"
    'SavePicture GFX.sprites.Picture, App.Path & "\temp\sprites.bmp"
    'SavePicture GFX.tiles.Picture, App.Path & "\temp\tiles.bmp"
    'SavePicture GFX.spells.Picture, App.Path & "\temp\spells.bmp"
    frmMirage.picItems.Picture = LoadPicture(App.Path & "\items.bmp")
    Form1.picItems.Picture = LoadPicture(App.Path & "\items.bmp")

    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\sprites.bmp", DDSD_Sprite)
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init tiles ddsd type and load the bitmap
    DDSD_Tile.lFlags = DDSD_CAPS
    DDSD_Tile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_TileSurf = DD.CreateSurfaceFromFile(App.Path & "\tiles.bmp", DDSD_Tile)
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\items.bmp", DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init spells ddsd type and load the bitmap
    DDSD_Spell.lFlags = DDSD_CAPS
    DDSD_Spell.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpellSurf = DD.CreateSurfaceFromFile(App.Path & "\spells.bmp", DDSD_Spell)
    DD_SpellSurf.SetColorKey DDCKEY_SRCBLT, key
    
    DDSD_Emote.lFlags = DDSD_CAPS
    DDSD_Emote.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
 Set DD_EmoteSurf = DD.CreateSurfaceFromFile(App.Path & "\emotes.bmp", DDSD_Emote)
    DD_EmoteSurf.SetColorKey DDCKEY_SRCBLT, key
    'buh bye ^_^
    'Kill (App.Path & "\temp\tiles.bmp")
    'Kill (App.Path & "\temp\sprites.bmp")
    'Kill (App.Path & "\temp\items.bmp")
    'Kill (App.Path & "\temp\spells.bmp")
    'Call RmDir(App.Path & "\temp")
    
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_SpellSurf = Nothing
    Set DD_EmoteSurf = Nothing
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        NeedToRestoreSurfaces = False
    Else
        NeedToRestoreSurfaces = True
    End If
End Function
