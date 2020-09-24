VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TerrainType
  Name As String
  SpriteName As String
  MoveCost As Byte
  Height As Byte
  Mask As Long
  Animated As Boolean
  Randomizable As Boolean
  FullSpritePath As String
  lWidth As Byte
  lHeight As Byte
  HideForInside As Boolean
  DXSurface As DirectDrawSurface7
End Type
Private Type ObjType
  X As Integer
  Y As Integer
  SpriteID As Integer
  InFront As Boolean 'drawn AFTER tiles
  HasHeight As Boolean 'is influenced by tile height?
End Type
Private Type ObjectSpriteType
  Surface As DirectDrawSurface7
  lWidth As Byte
  lHeight As Byte
  PathName As String
  SpriteRect As RECT
End Type

Public MapLoaded As Boolean

Public VisionX As Byte
Public VisionY As Byte

Public Graphics As Boolean

Public DefaultTileName As String

Public InsideView As Boolean
Public RoofLayer As Byte
Public DestObj As Object
Public LastDrawRatio As Single 'for draw code checking/debugging

Dim objDx As New DirectX7
Dim objDraw7 As DirectDraw7

Dim Terrain() As TerrainType

Dim Matrix() As Byte
Dim KeyMatrix() As Integer
Dim MyMaxDim As Integer
Dim MyMaxLayers As Byte
Dim MyTileWidth As Byte

Dim MyFocus As Vector2D

Dim Obj() As ObjType

Dim BackSurface As DirectDrawSurface7
Dim BackRect As RECT
Dim PrimarySurface As DirectDrawSurface7
Dim Clipper As DirectDrawClipper

Dim ObjectSpriteList() As ObjectSpriteType

Dim MyTextLine() As String
Dim MyLayerHidden() As Boolean

Dim DRDrawn As Long
Dim DRTotal As Long

'LayerHidden property allows you to set a certain layer to be hidden (= not drawn) or shown.
Public Property Get LayerHidden(Index As Byte) As Boolean
  LayerHidden = MyLayerHidden(Index)
End Property

Public Property Let LayerHidden(Index As Byte, Status As Boolean)
  MyLayerHidden(Index) = Status
End Property

'Clears all text output lines
Public Sub ClearTextLines()
  Dim n As Byte
  For n = 0 To UBound(MyTextLine)
    MyTextLine(n) = ""
  Next
End Sub

'TextLine property allows you to set text output lines that appear in the leftmost corner of the DestObj.
Public Property Get TextLine(ByVal ID As Byte) As String
  TextLine = MyTextLine(ID)
End Property

Public Property Let TextLine(ByVal ID As Byte, ByVal Value As String)
  If ID > UBound(MyTextLine) Then
    ReDim Preserve MyTextLine(ID)
  End If
  MyTextLine(ID) = Value
End Property

'Predicate that checks if a coordinate in the world is valid.
Public Function OnMap(X As Integer, Y As Integer, Optional L As Byte = 0) As Boolean
  OnMap = (X >= 0 And X <= MyMaxDim And Y >= 0 And Y <= MyMaxDim And L >= 0 And L <= MyMaxLayers)
End Function

'Sets the focus vector
Public Sub SetFocus(X As Integer, Y As Integer)
  FocusX = X
  FocusY = Y
End Sub

'X coordinate of focus vector
Public Property Get FocusX() As Integer
  FocusX = MyFocus.X
End Property

Public Property Let FocusX(Value As Integer)
'  MyFocusX = Cap(Value, VisionX, MaxDim)
  MyFocus.X = Value
End Property

'Y coordinate of focus vector
Public Property Get FocusY() As Integer
  FocusY = MyFocus.Y
End Property

Public Property Let FocusY(Value As Integer)
'  MyFocusY = Cap(Value, VisionY, MaxDim)
  MyFocus.Y = Value
End Property

'scrolls view in any given Direction
Public Sub Scroll(ByVal Direction As DirectionList)
  Dim NewFocus As Vector2D
  NewFocus = RefAdj(MyFocus, Direction)
  MyFocus.X = NewFocus.X
  MyFocus.Y = NewFocus.Y
End Sub

'abstractions for screen scrolling
Public Sub ScrollEast()
  Scroll dirEast
End Sub

Public Sub ScrollWest()
  Scroll dirWest
End Sub

Public Sub ScrollNorth()
  Scroll dirNorth
End Sub

Public Sub ScrollSouth()
  Scroll dirSouth
End Sub

'Width of a standard tile, used to compute drawing coordinates
Public Property Get TileWidth() As Byte
  TileWidth = MyTileWidth
End Property

Public Property Let TileWidth(Value As Byte)
  MyTileWidth = Value
End Property

'abstraction for editing a tile in the world matrix
Public Sub Edit(ByVal X As Integer, ByVal Y As Integer, ByVal L As Byte, ByVal State As Byte, Optional ByVal ShowErrors As Boolean = False)
  If OnMap(X, Y, L) Then
    If State >= 0 And State <= TerrainTypes Then
      Matrix(X, Y, L) = State
    Else
      If ShowErrors Then MsgBox "Terrain type " & State & " doesn't exist!", vbOKOnly + vbCritical, "Map Engine Error!"
    End If
  Else
    If ShowErrors Then MsgBox "Map reference errors! Point (" & X & ", " & Y & ", " & L & ") doesn't exist!", vbOKOnly + vbCritical, "Map Engine Error!"
  End If
End Sub

'loads the terrain from a datafile, also needs standard directory for tiles
Public Sub LoadTerrain(FileName As String, ByVal TileDir As String)
  Dim FID As Byte
  Dim FID2 As Byte
  Dim n As Byte
  Dim tStr As String
  Dim ddsTile As DDSURFACEDESC2
  Dim ddcKey As DDCOLORKEY
  Dim tLng As Long
  
  If Dir(FileName) <> "" Then 'if file exists
    FID = FreeFile
    Open FileName For Input As #FID
      n = 0 'set counter to zero
      While Not EOF(FID) 'load as long as the file contains text
        n = n + 1
        ReDim Preserve Terrain(n) 'enlarge terrain array
        With Terrain(n)
          Input #FID, tStr 'read line
          .SpriteName = tStr 'first line is sprite filename
          
          If Graphics Then 'if this instance of the engine is graphical
            .FullSpritePath = TileDir & .SpriteName & ".bmp" 'full path of the sprite
            FID2 = FreeFile
            ddsTile.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
            'this is for determining the dimensions of the bitmap
            Open .FullSpritePath For Binary Access Read As #FID2 'OPEN THE BITMAP AS BINARY
              Get #FID2, 19, tLng: ddsTile.lWidth = tLng '19th DOUBLE WORD is Width
              Get #FID2, 23, tLng: ddsTile.lHeight = tLng '23rd DOUBLE WORD is Height
            Close #FID2
            .lWidth = ddsTile.lWidth
            .lHeight = ddsTile.lHeight
            'to load and init the surface
            ddsTile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
            Set Terrain(n).DXSurface = objDraw7.CreateSurfaceFromFile(.FullSpritePath, ddsTile) 'create a surface from the file
            .DXSurface.SetColorKey DDCKEY_SRCBLT, ddcKey 'set transparancy (black is standard set)
          End If
          'other props
          Input #FID, tStr: .Name = tStr 'read name (= category) from file
          Input #FID, tStr: .MoveCost = tStr
          Input #FID, tStr: .Height = tStr
          Input #FID, tStr: .Animated = CBool(tStr)
          Input #FID, tStr: .Randomizable = CBool(tStr)
          Input #FID, tStr: .HideForInside = CBool(tStr)
        End With
      Wend
    Close #FID
  Else
    MsgBox "File " & FileName & " not found!", vbCritical + vbOKOnly, "File not found!"
  End If
End Sub

'returns the rectangle defined by the dimensions of a certain tile
Private Function TerrainRect(TileID As Byte) As RECT
  TerrainRect.Left = 0
  TerrainRect.Top = 0
  TerrainRect.Right = Terrain(TileID).lWidth
  TerrainRect.Bottom = Terrain(TileID).lHeight
End Function

'count of terraintypes
Public Property Get TerrainTypes() As Byte
  TerrainTypes = UBound(Terrain)
End Property

'maximum dimension of the (square) world Matrix
'I might change this to non-square worlds
Public Property Let MaxDim(Value As Integer)
  If Value = MyMaxDim Then Exit Property
  MyMaxDim = Value
  ReDim Preserve Matrix(0 To Value, 0 To Value, 0 To MyMaxLayers)
  ReDim Preserve KeyMatrix(0 To Value, 0 To Value)
End Property

Public Property Get MaxDim() As Integer
  MaxDim = MyMaxDim
End Property

'layers 0 to MaxLayers
Public Property Let MaxLayers(Value As Integer)
  If Value = MyMaxLayers Then Exit Property
  MyMaxLayers = Value
  ReDim Preserve Matrix(0 To MyMaxDim, 0 To MyMaxDim, 0 To Value)
  ReDim MyLayerHidden(0 To Value)
End Property

Public Property Get MaxLayers() As Integer
  MaxLayers = MyMaxLayers
End Property

'initialises everything that's got to do with DirectX
Public Sub InitDirectX(Optional SwitchRes As Boolean = True)
  Dim ddsdScreen As DDSURFACEDESC2 'foreground
  Dim ddsdBack As DDSURFACEDESC2 'backbuffer
  
  Graphics = True 'graphics are activated
  
  Set objDraw7 = objDx.DirectDrawCreate("") 'init DirectDraw7
  objDraw7.SetCooperativeLevel DestObj.hwnd, DDSCL_NORMAL 'set it to work with the destination form
  If SwitchRes Then objDraw7.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT '(maybe) switch resolution
  
  ddsdScreen.lFlags = DDSD_CAPS
  ddsdScreen.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
  
  Set PrimarySurface = objDraw7.CreateSurface(ddsdScreen) 'create primary surface
  
  Set Clipper = objDraw7.CreateClipper(0) 'create clipper to bound a region to the screen
  
  Clipper.SetHWnd DestObj.hwnd
  
  PrimarySurface.SetClipper Clipper
  
  ddsdBack.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
  ddsdBack.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  ddsdBack.lWidth = DestObj.ScaleWidth
  ddsdBack.lHeight = DestObj.ScaleHeight
  
  Set BackSurface = objDraw7.CreateSurface(ddsdBack) 'create backbuffer surface
  
  'set backbuffer rect
  BackRect.Bottom = ddsdBack.lHeight
  BackRect.Right = ddsdBack.lWidth
  
  'make it look blue to show the "void tiles"
  BackSurface.SetFillColor vbBlue
End Sub

'dimension abstractions of the dest object
Public Property Get VideoHeight() As Integer
  VideoHeight = DestObj.ScaleHeight
End Property

Public Property Get VideoWidth() As Integer
  VideoWidth = DestObj.ScaleWidth
End Property

'draw the world to the DestObj object
Public Sub DrawMap()
  Dim X As Integer
  Dim Y As Integer
  Dim sX As Integer
  Dim sY As Integer
  Dim eX As Integer
  Dim eY As Integer
  Dim C As Integer
  Dim n As Integer
  Dim tRect As RECT
  Dim DrawPos As Vector2D
  Dim Draw As Boolean
  
  'clear the back surface, not necessary but recommended
  BackSurface.DrawBox 0, 0, VideoWidth, VideoHeight
  
  'these are little counters to check what percentage of the computed area is actually drawn
  DRTotal = 0
  DRDrawn = 0
  
  'get bounding boxes of "visible area"
  'should really be optimized!
  sX = LeftX
  sY = LeftY
  eX = RightX
  eY = RightY
  For Y = sY To eY
    For X = sX To eX
      DrawTile X, Y 'draw each "visible" tile
    Next
  Next
  
  If DRTotal > 0 Then
    LastDrawRatio = DRDrawn / DRTotal 'compute drawn ratio
  End If
  
  'draws all game objects that are InFront, like cursors and the likes
  C = UBound(Obj)
  For n = 1 To C
    With Obj(n)
      If .InFront Then
        'get the location
        DrawPos = GetObjectDrawCoords(.X, .Y, ObjectSpriteList(.SpriteID).lWidth, ObjectSpriteList(.SpriteID).lHeight)
        'get the draw rect
        tRect = ObjectSpriteList(.SpriteID).SpriteRect
        'confine the location and rect to fit viewscreen (to avoid unnecessary blitting)
        'the boolean return value is false if the sprite is entirely off screen
        Draw = ConfineCoordinates(DrawPos, tRect)
        If Draw Then
          'draw to back surface
          BackSurface.BltFast DrawPos.X, DrawPos.Y, ObjectSpriteList(.SpriteID).Surface, tRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
      End If
    End With
  Next
  
  'drawing text lines
  'these lines are split on the first : character because it contains chat text
  '
  '  Joey: Hey how you doin'?
  '
  'is treated differently from
  '
  '  Joey asks how I am doin'.
  '
  Const Color1 = &H40D0D0
  Const Color2 = &H80C0F0
  Const Color3 = &HA0C0D0
  Dim tX As Integer
  Dim tY As Integer
  Dim Title As String
  Dim Text As String
  Dim i As Byte
  Const PosX = 20
  Const PosY = 10
  
  For n = 0 To UBound(MyTextLine)
    If MyTextLine(n) <> "" Then
      tX = 12 + PosX
      tY = 24 + (n - 1) * 15 + PosY
      i = InStr(MyTextLine(n), ":")
      If i = 0 Then
        BackSurface.SetForeColor Color2
        BackSurface.DrawText tX, tY, MyTextLine(n), False
      Else
        Title = Left(MyTextLine(n), i - 1)
        Text = Right(MyTextLine(n), Len(MyTextLine(n)) - i)
        BackSurface.SetForeColor Color1
        BackSurface.DrawText tX, tY, Title, False
        BackSurface.SetForeColor Color3
        tX = 12 + 150 + PosX
        tY = 24 + (n - 1) * 15 + PosY
        BackSurface.DrawText tX, tY, Text, False
      End If
    End If
  Next
  
  'Get a rect or something to blit to
  objDx.GetWindowRect DestObj.hwnd, tRect
  'blit the whole backbuffer to primary surface (the actual refreshing)
  PrimarySurface.Blt tRect, BackSurface, BackRect, DDBLT_WAIT
End Sub

'this code makes it possible for locations that lie outside the clipper to be drawn to screen
'the leftmost point is shifted so it still fits in screen, while the Rect is changed as well
'if the tile falls outside the clipper alltogether, False s returned
Private Function ConfineCoordinates(VectPos As Vector2D, DrawRect As RECT) As Boolean
  ConfineCoordinates = True
  If VectPos.X < 0 Then
    If VectPos.X + DrawRect.Right > 0 Then
      DrawRect.Left = -VectPos.X
      VectPos.X = 0
    Else
      ConfineCoordinates = False
    End If
  ElseIf VectPos.X + DrawRect.Right > VideoWidth Then
    If VectPos.X < VideoWidth Then
      DrawRect.Right = VideoWidth - VectPos.X
    Else
      ConfineCoordinates = False
    End If
  End If
  If VectPos.Y < 0 Then
    If VectPos.Y + DrawRect.Bottom > 0 Then
      DrawRect.Top = -VectPos.Y
      VectPos.Y = 0
    Else
      ConfineCoordinates = False
    End If
  ElseIf VectPos.Y + DrawRect.Bottom > VideoHeight Then
    If VectPos.Y < VideoHeight Then
      DrawRect.Bottom = VideoHeight - VectPos.Y
    Else
      ConfineCoordinates = False
    End If
  End If
End Function

'draws all tiles and objects that are located at X, Y
Public Sub DrawTile(X As Integer, Y As Integer)
  Dim DrawPos As Vector2D
  Dim T As Byte
  Dim L As Byte
  Dim n As Integer
  Dim RoofTile As Byte
  Dim tRect As RECT
  Dim Draw As Boolean
  Dim H As Integer
  
  'checks the location
  If Not OnMap(X, Y) Then Exit Sub
  
  'count drawing status
  DRTotal = DRTotal + 1
  
  'IF the primary layer contains no tile, the whole location is ignored (Gap In Reality, as if)
  If Matrix(X, Y, 0) = 0 Then Exit Sub
  
  'draws layer 0 and 1, the primary ground layer and top surface layer (grass and trees, floor and tables, ...)
  For L = 0 To OfMax(1, MaxLayers)
    If Not LayerHidden(L) Then
      T = Matrix(X, Y, L)
      If T > 0 Then
        'get coords
        DrawPos = GetTileDrawCoords(X, Y, L)
        'get rect
        tRect = TerrainRect(T)
        'confine
        Draw = ConfineCoordinates(DrawPos, tRect)
        'if confined in clipper, draw
        If Draw Then
          'blit to backbuffer
          BackSurface.BltFast DrawPos.X, DrawPos.Y, Terrain(T).DXSurface, tRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If L = 0 Then
            'count as drawn
            DRDrawn = DRDrawn + 1
          End If
        End If
      End If
    End If
  Next

  'compute the height of the terrain for Game Objects that are affected by height (players, items, yada yada)
  H = TerrainHeight(X, Y)
  L = UBound(Obj)
  For n = 1 To L 'loop through all game objects
    With Obj(n)
      If Not .InFront Then 'don't consider those that are InFront (cursors, etc...)
        'all these checks are seperate, as optimization, because VB doesn't break in AND combinators :-(
        If .SpriteID > 0 Then 'if valid sprite is assigned
          If .X = X Then 'check X
            If .Y = Y Then 'check Y
              'compute position
              DrawPos = GetObjectDrawCoords(X, Y, ObjectSpriteList(.SpriteID).lWidth, ObjectSpriteList(.SpriteID).lHeight)
              'if affected by height, change that position accordingly
              If .HasHeight Then DrawPos.Y = DrawPos.Y - H
              'get rect
              tRect = ObjectSpriteList(.SpriteID).SpriteRect
              'confine
              Draw = ConfineCoordinates(DrawPos, tRect)
              'if in clipper...
              If Draw Then
                'blit to backsurface
                BackSurface.BltFast DrawPos.X, DrawPos.Y, ObjectSpriteList(.SpriteID).Surface, tRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
              End If
            End If
          End If
        End If
      End If
    End With
  Next
  
  If MaxLayers <= 1 Then Exit Sub
  
  '"ROOF" system for all layers > 1
  RoofTile = 0 'the "kind" of roofing
  If InsideView Then 'if the focus is affected by roofs
    RoofTile = Matrix(FocusX, FocusY, RoofLayer) 'assign roof tile
  End If
  For L = 2 To MaxLayers 'loop through remaining layers
    If Not LayerHidden(L) Then
      T = Matrix(X, Y, L)
      If T > 0 Then
        'if the tile isn't the same roof as the focus is under or the focus is not under a roof, then draw...
        If Not (L = RoofLayer And Terrain(T).Name = Terrain(RoofTile).Name And Terrain(T).HideForInside) Or RoofTile = 0 Then
          'get loc
          DrawPos = GetTileDrawCoords(X, Y, L)
          'get rect
          tRect = TerrainRect(T)
          'confine
          Draw = ConfineCoordinates(DrawPos, tRect)
          If Draw Then
            'draw to backbuffer
            BackSurface.BltFast DrawPos.X, DrawPos.Y, Terrain(T).DXSurface, tRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          End If
        End If
      End If
    End If
  Next
End Sub

'maximum height found on a location
Public Function TerrainHeight(ByVal X As Integer, ByVal Y As Integer) As Integer
  Dim L As Byte
  Dim T As Byte
  For L = 0 To MaxLayers
    T = Matrix(X, Y, L)
    If T > 0 Then
      TerrainHeight = Max(TerrainHeight, Terrain(T).Height)
    End If
  Next
End Function

'edge of the maximum visible area
'these need to be optimized, because a lot of area is taken into consideration, because the screen
'is an orthogonal projection on a "diagonal non-orthogonal" axis sytem
Public Property Get LeftX() As Integer
  LeftX = OfMin(MyFocus.X - VisionX, 0)
End Property

Public Property Get LeftY() As Integer
  LeftY = OfMin(MyFocus.Y - VisionY, 0)
End Property

Public Property Get RightX() As Integer
  RightX = OfMax(MyFocus.X + VisionX, MaxDim)
End Property

Public Property Get RightY() As Integer
  RightY = OfMax(MyFocus.Y + VisionY, MaxDim)
End Property

'the location where a tile is to be drawn, abstraction to General Object
Private Function GetTileDrawCoords(ByVal X As Integer, ByVal Y As Integer, ByVal Layer As Byte) As Vector2D
  GetTileDrawCoords = GetObjectDrawCoords(X, Y, Terrain(Matrix(X, Y, Layer)).lWidth, Terrain(Matrix(X, Y, Layer)).lHeight)
End Function

'This is the heart of the Graphics Engine. Good luck!
Private Function GetObjectDrawCoords(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Byte, ByVal Height As Byte) As Vector2D
  'get relative coords. Is allowed this way because X and Y are ByVal
  X = X - MyFocus.X
  Y = Y - MyFocus.Y
  'have a deep breath and look at this. then get coffee
  GetObjectDrawCoords.X = VideoWidth \ 2 + (Y - X) * TileWidth \ 2 - Width \ 2
  GetObjectDrawCoords.Y = VideoHeight \ 2 + (Y + X) * TileWidth \ 4 + TileWidth \ 2 - 1 - Height
End Function

'returns the tile coordinates that a possible mouse cursor or the such is pointing to. Very helpfull!
Public Function GetTileX(ByVal PixelX As Long, ByVal PixelY As Long) As Integer
  GetTileX = ((PixelY - VideoHeight \ 2) / (TileWidth \ 4) - (PixelX - VideoWidth \ 2 + TileWidth \ 2) / (TileWidth \ 2)) / 2 + FocusX
End Function

Public Function GetTileY(ByVal PixelX As Long, ByVal PixelY As Long) As Integer
  GetTileY = ((PixelX - VideoWidth \ 2 + TileWidth \ 2) / (TileWidth \ 2) + (PixelY - VideoHeight \ 2) / (TileWidth \ 4)) / 2 + FocusY
End Function

'all randomizable class tiles are randomized to random value in their category
Public Sub RandomizeTiles()
  Dim X As Integer
  Dim Y As Integer
  Dim L As Byte
  Dim tAr() As Byte
  'default tile list is a heuristic to keep the most frequently used tile (default tile) in short memory
  'sort of a limited memoize technique
  Dim DefaultTL() As Byte
  Dim DefaultReady As Boolean
  
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      For L = 0 To MaxLayers
        If Terrain(Matrix(X, Y, L)).Randomizable Then
          If Terrain(Matrix(X, Y, L)).Name = DefaultTileName Then
            If Not DefaultReady Then
              DefaultTL = AssembleTerrainList(Terrain(Matrix(X, Y, L)).Name)
              DefaultReady = True
            End If
            tAr = DefaultTL
          Else
            tAr = AssembleTerrainList(Terrain(Matrix(X, Y, L)).Name)
          End If
          Matrix(X, Y, L) = tAr(Rand(0, UBound(tAr)))
        End If
      Next
    Next
  Next
End Sub

'fill the whole world by Name of a tile (category)
Public Sub FillByName(Name As String)
  Dim X As Integer
  Dim Y As Integer
  Dim tAr() As Byte
  If Name = "" Then Exit Sub
  tAr = AssembleTerrainList(Name)
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      Matrix(X, Y, 0) = tAr(Rand(0, UBound(tAr))) 'randomized!
    Next
  Next
End Sub

'fill by tile ID
Public Sub Fill(TileID As Byte)
  Dim X As Integer
  Dim Y As Integer
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      Matrix(X, Y, 0) = TileID
    Next
  Next
End Sub

'wipes off everything on every layer, sets it all to Void
Public Sub Clear()
  Dim X As Integer
  Dim Y As Integer
  Dim L As Byte
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      For L = 0 To MaxLayers
        Matrix(X, Y, L) = 0
      Next
    Next
  Next
End Sub

'clear the game object matrix
Public Sub ClearObjects()
  ReDim Obj(0)
End Sub

'total number of objects
Public Property Get ObjectCount() As Integer
  ObjectCount = UBound(Obj)
End Property

'Add an object. Objects are dynamic things in the Engine, such as cursors, players, items, etc...
'The path to their sprite is assigned and then the tile is loaded and placed into an array.
'The loading only happens once each session, because an extra array is used to remember previous tiles.
Public Function AddObject(FullSpritePathName As String, ByVal X As Integer, ByVal Y As Integer, Optional InFront As Boolean = False, Optional ByVal HasHeight As Boolean = False) As Integer
  Dim L As Integer
  If Not OnMap(X, Y) Then Exit Function
  L = UBound(Obj)
  ReDim Preserve Obj(L + 1)
  With Obj(L + 1)
    .SpriteID = GetObjectSpriteID(FullSpritePathName)
    .X = X
    .Y = Y
    .InFront = InFront
    .HasHeight = HasHeight
  End With
  AddObject = L + 1 'return object count
End Function

'returns the number of cached sprites. cached sprites are sprites that have already once been loaded
Public Property Get CachedSprites() As Integer
  CachedSprites = UBound(ObjectSpriteList)
End Property

'this function does two things
'for new tiles, it loads them and assigns an ID in its matrix
'for already cached tiles, it looks up their ID
Private Function GetObjectSpriteID(PathName As String) As Integer
  Dim n As Integer
  Dim L As Integer
  Dim FID As Byte
  Dim tLng As Long
  Dim ddsTile As DDSURFACEDESC2
  Dim ddcKey As DDCOLORKEY
  
  If Dir(PathName) = "" Then Exit Function
  
  'first try looking up the ID
  L = UBound(ObjectSpriteList)
  For n = 1 To L
    If StrComp(ObjectSpriteList(n).PathName, PathName, vbTextCompare) = 0 Then
      GetObjectSpriteID = n
      Exit Function
    End If
  Next
  
  'if it aint there, load it from file
  ReDim Preserve ObjectSpriteList(L + 1)
  With ObjectSpriteList(L + 1)
    'open file as Binary to get dimensions
    FID = FreeFile
    Open PathName For Binary Access Read As #FID
      Get #FID, 19, tLng: ddsTile.lWidth = tLng '19th Double Word
      Get #FID, 23, tLng: ddsTile.lHeight = tLng '23rd Double Word
    Close #FID
    .lWidth = ddsTile.lWidth
    .lHeight = ddsTile.lHeight
    
    'assign rect
    .SpriteRect.Left = 0
    .SpriteRect.Top = 0
    .SpriteRect.Right = .lWidth
    .SpriteRect.Bottom = .lHeight
    
    .PathName = PathName
    'load the surface
    ddsTile.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsTile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set .Surface = objDraw7.CreateSurfaceFromFile(PathName, ddsTile) 'load surface from File!
    .Surface.SetColorKey DDCKEY_SRCBLT, ddcKey
  End With
  GetObjectSpriteID = L + 1 'return the new index
End Function

'returns the array of all terrain ID's of a certain name (= category)
Public Function AssembleTerrainList(Name As String) As Byte()
  Dim n As Byte
  Dim C As Byte
  Dim tAr() As Byte
  If Name = "" Then Exit Function
  C = 0
  For n = 1 To TerrainTypes
    If Terrain(n).Name = Name Then
      ReDim Preserve tAr(C)
      tAr(C) = n
      C = C + 1
    End If
  Next
  AssembleTerrainList = tAr
End Function

'returns an array of every name in the terrain matrix (each name only once, duh)
Public Function TerrainNameList() As String()
  Dim n As Byte
  Dim i As Integer
  Dim tAr() As String
  Dim tStr As String
  Dim L As Integer
  ReDim tAr(0)
  For n = 1 To TerrainTypes
    tStr = Terrain(n).Name
    L = OfMax(n - 1, UBound(tAr))
    For i = 1 To L
      If tStr = tAr(i) Then Exit For
    Next
    If i = L + 1 Then
      L = UBound(tAr) + 1
      If L > 0 Then ReDim Preserve tAr(L)
      tAr(L) = tStr
    End If
  Next
  TerrainNameList = tAr
End Function

'the world animator. Animated tiles are animated by this AnimateTiles sub :-)
'for water, or fires
Public Sub AnimateTiles()
  Dim X As Integer
  Dim Y As Integer
  Dim tAr() As Byte
  Dim L As Byte
  For X = LeftX To RightX
    For Y = LeftY To RightY
      If OnMap(X, Y) Then
        For L = 0 To MaxLayers
          With Terrain(Matrix(X, Y, L))
            If .Animated Then
              tAr() = AssembleTerrainList(.Name)
              Matrix(X, Y, L) = tAr(Rand(0, UBound(tAr)))
            End If
          End With
        Next
      End If
    Next
  Next
End Sub

'saves the entire map matrix to file
'several methods are known, the latter have best compression but there is always backwards compatibility
'I will not comment this because I don't want to give away all my cards
Public Sub SaveMap(FileName As String, Optional SaveMethod As Byte = 3)
  Dim X As Integer
  Dim Y As Integer
  Dim L As Byte
  Dim FID As Byte
  Dim T As Byte
  Dim B1 As Byte
  Dim B2 As Byte
  Dim B3 As Byte
  Dim KeyStr As String
  FID = FreeFile
  If Dir(FileName) <> "" Then Kill FileName
  Open FileName For Binary Access Write As #FID
    Put #FID, , CByte(SaveMethod)
    Put #FID, , CInt(MaxDim)
    Put #FID, , CByte(MaxLayers)
    Select Case SaveMethod
      
      Case 1
        For Y = 0 To MaxDim
          For X = 0 To MaxDim
            For L = 0 To MaxLayers
              T = Matrix(X, Y, L)
              If T > 0 Then
                If Terrain(T).Name <> DefaultTileName Then
                  Put #FID, , CInt(X)
                  Put #FID, , CInt(Y)
                  Put #FID, , CByte(L)
                  Put #FID, , CByte(T)
                  Put #FID, , CInt(KeyMatrix(X, Y))
                End If
              End If
            Next
          Next
        Next
      
      Case 2
        For Y = 0 To MaxDim
          For X = 0 To MaxDim
            For L = 0 To MaxLayers
              T = Matrix(X, Y, L)
              If T > 0 And T <= TerrainTypes Then
                If Terrain(T).Name <> DefaultTileName Then
                  B1 = LSB(X)
                  B2 = LSB(Y)
                  B3 = MSB(X) + MSB(Y) * 16
                  Put #FID, , CByte(B1)
                  Put #FID, , CByte(B2)
                  Put #FID, , CByte(B3)
                  Put #FID, , CByte(L)
                  Put #FID, , CByte(T)
                  Put #FID, , CInt(KeyMatrix(X, Y))
                End If
              End If
            Next
          Next
        Next
      
      Case 3
        KeyStr = ""
        For Y = 0 To MaxDim
          For X = 0 To MaxDim
            If KeyMatrix(X, Y) > 0 Then
              B1 = LSB(X)
              B2 = LSB(Y)
              B3 = MSB(X) + MSB(Y) * 16
              KeyStr = KeyStr & Chr(B1) & Chr(B2) & Chr(B3) & Chr(LSB(KeyMatrix(X, Y))) & Chr(MSB(KeyMatrix(X, Y)))
            End If
          Next
        Next
        Put #FID, , CLng(Len(KeyStr))
        Put #FID, , KeyStr
        For Y = 0 To MaxDim
          For X = 0 To MaxDim
            For L = 0 To MaxLayers
              T = Matrix(X, Y, L)
              If T > 0 And T <= UBound(Terrain) Then
                If Terrain(T).Name <> DefaultTileName Then
                  B1 = LSB(X)
                  B2 = LSB(Y)
                  B3 = MSB(X) + MSB(Y) * 16
                  Put #FID, , CByte(B1)
                  Put #FID, , CByte(B2)
                  Put #FID, , CByte(B3)
                  Put #FID, , CByte(L)
                  Put #FID, , CByte(T)
                End If
              End If
            Next
          Next
        Next
    
      Case Else
        MsgBox "Unknown MAPMODE " & SaveMethod & "!", vbOKOnly + vbCritical, "Error in SaveMap!"
    
    End Select
  Close #FID
End Sub

'loads the entire map matrix to file
'several methods are known, the latter have best compression but there is always backwards compatibility
'I will not comment this because I don't want to give away all my cards
Public Sub LoadMap(FileName As String)
  Dim X As Integer
  Dim Y As Integer
  Dim L As Byte
  Dim FID As Byte
  Dim tByte As Byte
  Dim tInt As Integer
  Dim SaveMethod As Byte
  Dim KeyLen As Long
  Dim B1 As Byte
  Dim B2 As Byte
  Dim B3 As Byte
  Clear
  FillByName DefaultTileName
  FID = FreeFile
  If Dir(FileName) = "" Then Exit Sub
  Open FileName For Binary Access Read As #FID
    '//HEADER//
    Get #FID, , tByte: SaveMethod = tByte 'saving method
    Get #FID, , tInt ': MaxDim = tInt
    Get #FID, , tByte ': MaxLayers = tByte
    '//BODY//
    Select Case SaveMethod
        
      Case 1
        While Not EOF(FID)
          Get #FID, , X
          Get #FID, , Y
          Get #FID, , L
          Get #FID, , tByte: Matrix(X, Y, L) = tByte
          Get #FID, , tInt: KeyMatrix(X, Y) = tInt
        Wend
        
      Case 2
        While Not EOF(FID)
          Get #FID, , B1
          Get #FID, , B2
          Get #FID, , B3
          X = B1 + (B3 And 15) * 256
          Y = B2 + (B3 And 240) * 16
          Get #FID, , L
          Get #FID, , tByte: Matrix(X, Y, L) = tByte
          Get #FID, , tInt: KeyMatrix(X, Y) = tInt
        Wend
        
      Case 3
        Get #FID, , KeyLen
        While KeyLen > 0
          Get #FID, , B1
          Get #FID, , B2
          Get #FID, , B3
          X = B1 + (B3 And 15) * 256
          Y = B2 + (B3 And 240) * 16
          Get #FID, , tInt
          KeyMatrix(X, Y) = tInt
          KeyLen = KeyLen - 5
        Wend
        While Not EOF(FID)
          Get #FID, , B1
          Get #FID, , B2
          Get #FID, , B3
          X = B1 + (B3 And 15) * 256
          Y = B2 + (B3 And 240) * 16
          Get #FID, , L
          Get #FID, , tByte: Matrix(X, Y, L) = tByte
        Wend
      
      Case Else
        MsgBox "Unknown MAPMODE " & SaveMethod & "!", vbOKOnly + vbCritical, "Error in LoadMap!"
      
    End Select
  Close #FID
  MapLoaded = True
End Sub

'returns the Index of any terrain by Name
Public Function GetTerrainID(Name As String) As Byte
  Dim n As Byte
  For n = 1 To UBound(Terrain)
    If StrComp(LCase(Terrain(n).Name), Name, vbTextCompare) = 0 Then
      GetTerrainID = n
      Exit Function
    End If
  Next
End Function

'returns the Index of any terrain by Sprite Name
Public Function GetTerrainIDBySprite(SpriteName As String) As Byte
  Dim n As Byte
  For n = 1 To UBound(Terrain)
    If StrComp(LCase(Terrain(n).SpriteName), SpriteName, vbTextCompare) = 0 Then
      GetTerrainIDBySprite = n
      Exit Function
    End If
  Next
End Function

'Map Key at a location. Map keys are to be used for, you guessed it, doors :-)
Public Property Get KeyAt(ByVal X As Integer, ByVal Y As Integer) As Integer
  KeyAt = KeyMatrix(X, Y)
End Property

Public Property Let KeyAt(ByVal X As Integer, ByVal Y As Integer, ByVal KEYVAL As Integer)
  KeyMatrix(X, Y) = KEYVAL
End Property

'This code is tough. it's used to clean up the edges of water. Look into it if you wanna
Public Sub FixBeachOnTile(X As Integer, Y As Integer)
  Dim FLAG As Byte
  Dim n As Byte
  Dim sP As Vector2D
  Dim oP As Vector2D
  If Not OnMap(X, Y) Then Exit Sub
  If Terrain(Matrix(X, Y, 0)).Name = "water" Then Exit Sub
  FLAG = 0
  sP = ToVector(X, Y)
  For n = 1 To 7 Step 2
    oP = RefAdj(sP, n)
    If OnMap(oP.X, oP.Y) Then
      If Terrain(Matrix(oP.X, oP.Y, 0)).Name = "water" Then
        FLAG = (FLAG Or 2 ^ ((n - 1) * 0.5))
      End If
    End If
  Next
  If FLAG > 0 Then
    Matrix(X, Y, 1) = GetTerrainID("beach" & FLAG)
  End If
End Sub

'call this sub to round the edges on water tiles. this is all done to layer 1
Public Sub FixBeaches()
  Dim X As Integer
  Dim Y As Integer
  Dim T As Integer
  Dim n As Byte
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      With Terrain(Matrix(X, Y, 0))
        If .Name <> "water" Then
          If Matrix(X, Y, 1) > 0 Then
            If Terrain(Matrix(X, Y, 1)).Name = "beach" Then
              FixBeachOnTile X, Y
            End If
          Else
            FixBeachOnTile X, Y
          End If
        End If
      End With
    Next
  Next
End Sub

'Checks neighbouring tiles for a road on a given location and fixes them accordingly
Public Sub FixRoadOnTile(X As Integer, Y As Integer)
  Dim FLAG As Byte
  Dim n As Byte
  Dim sP As Vector2D
  Dim oP As Vector2D
  If Not OnMap(X, Y) Then Exit Sub
  If Left(Terrain(Matrix(X, Y, 1)).Name, 4) <> "road" Then Exit Sub
  FLAG = 0
  sP = ToVector(X, Y)
  For n = 1 To 7 Step 2
    oP = RefAdj(sP, n)
    If OnMap(oP.X, oP.Y) Then
      If Left(Terrain(Matrix(oP.X, oP.Y, 1)).Name, 4) = "road" Then
        FLAG = (FLAG Or 2 ^ ((n - 1) * 0.5))
      End If
    End If
  Next
  If FLAG > 0 Then
    Matrix(X, Y, 1) = GetTerrainID("road" & FLAG)
  End If
End Sub

'fixes all roads on the whole map
Public Sub FixRoads()
  Dim X As Integer
  Dim Y As Integer
  Dim T As Integer
  Dim n As Byte
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      If Matrix(X, Y, 1) > 0 Then
        If Left(Terrain(Matrix(X, Y, 1)).Name, 4) = "road" Then
          FixRoadOnTile X, Y
        End If
      End If
    Next
  Next
End Sub

'Returns the Terrain ID at a location. Kinda obvious
Public Property Get Tile(X As Integer, Y As Integer, L As Byte) As Byte
  If OnMap(X, Y, L) Then
    Tile = Matrix(X, Y, L)
  End If
End Property

'this one is pretty cool.
'you input a location and a facing direction and it returns a vector to the "new" location one tile in that direction
Private Function RefAdj(LOC As Vector2D, ByVal Direction As DirectionList) As Vector2D
  Select Case Direction
    Case 0 'N
      RefAdj.X = LOC.X - 1
      RefAdj.Y = LOC.Y - 1
    Case 1 'NE
      RefAdj.X = LOC.X - 1
      RefAdj.Y = LOC.Y
    Case 2 'E
      RefAdj.X = LOC.X - 1
      RefAdj.Y = LOC.Y + 1
    Case 3 'SE
      RefAdj.X = LOC.X
      RefAdj.Y = LOC.Y + 1
    Case 4 'S
      RefAdj.X = LOC.X + 1
      RefAdj.Y = LOC.Y + 1
    Case 5 'SW
      RefAdj.X = LOC.X + 1
      RefAdj.Y = LOC.Y
    Case 6 'W
      RefAdj.X = LOC.X + 1
      RefAdj.Y = LOC.Y - 1
    Case 7 'NW
      RefAdj.X = LOC.X
      RefAdj.Y = LOC.Y - 1
  End Select
End Function

'abstraction for name of a terrain ID
Public Property Get TileSpriteName(ID As Byte) As String
  TileSpriteName = Terrain(ID).SpriteName
End Property

'inits objects and cache list
Public Sub Init()
  ClearObjects
  ReDim ObjectSpriteList(0)
End Sub

'obvious
Public Sub UnloadDirectX()
  objDraw7.RestoreDisplayMode
End Sub

'for drawing bounding. this has to change
Public Property Get InViewRange(X As Integer, Y As Integer)
  InViewRange = (X >= LeftX And X <= RightX And Y >= LeftY And Y <= RightY)
End Property

'relative version of latter
Public Property Get InViewRangeForXY(X As Integer, Y As Integer, OriginX As Integer, OriginY As Integer)
  InViewRangeForXY = (X >= OriginX - VisionX And X <= OriginX + VisionX And Y >= OriginY - VisionY And Y <= OriginY + VisionY)
End Property

'returns the movecost at a certain location
'this is rather complex because it needs to take the max at each layer
Public Property Get MoveCostAt(X As Integer, Y As Integer)
  Dim L As Byte
  If Not MapLoaded Then Exit Property
  If Not OnMap(X, Y) Then Exit Property
  MoveCostAt = Terrain(Matrix(X, Y, 0)).MoveCost
  If MoveCostAt = 0 Then Exit Property
  For L = 1 To MyMaxLayers
    If Matrix(X, Y, L) > 0 Then
      If Terrain(Matrix(X, Y, L)).MoveCost = 0 Then
        MoveCostAt = 0
        Exit Property
      Else
        MoveCostAt = Max(MoveCostAt, Terrain(Matrix(X, Y, L)).MoveCost)
      End If
    End If
  Next
End Property

'a tile is blocking if it has movecost zero
Public Property Get BlockingAt(X As Integer, Y As Integer)
  BlockingAt = (MoveCostAt(X, Y) = 0)
End Property

'toggles all doors at a given location
Public Property Get ToggleDoorsAt(X As Integer, Y As Integer) As Byte
  If Not MapLoaded Then Exit Property
  If Not OnMap(X, Y) Then Exit Property
  Dim L As Byte
  Dim T As Byte
  For L = 0 To MaxLayers
    T = Matrix(X, Y, L)
    If T > 0 Then
      If Terrain(T).Name = "door" Then
        If Left(Terrain(T).SpriteName, 4) = "open" Then
          'an open door: close it!
          Matrix(X, Y, L) = GetTerrainIDBySprite(Replace(Terrain(T).SpriteName, "open", "closed"))
          ToggleDoorsAt = 2
        Else
          'a closed door: open it!
          Matrix(X, Y, L) = GetTerrainIDBySprite(Replace(Terrain(T).SpriteName, "closed", "open"))
          ToggleDoorsAt = 1
        End If
        Exit Property
      End If
    End If
  Next
  ToggleDoorsAt = 0
End Property

'name of terrain at X,Y
Public Property Get TerrainNameAt(X As Integer, Y As Integer) As String
  TerrainNameAt = Terrain(Matrix(X, Y, 0)).Name
End Property

'checks if any tile near the given coords is of name Name
Public Property Get IsNearTile(X As Integer, Y As Integer, Name As String) As Boolean
  Dim n As Byte
  Dim vPos As Vector2D
  Dim tVect As Vector2D
  Dim L As Byte
  vPos = ToVector(X, Y)
  For n = 0 To 7
    tVect = RefAdj(vPos, n)
    For L = 0 To MaxLayers
      If Terrain(Matrix(tVect.X, tVect.Y, L)).Name = Name Then
        IsNearTile = True
        Exit Property
      End If
    Next L
  Next
End Property

'returns if the tile at a location is of category/name "door"
Public Property Get IsDoorAt(ByVal X As Integer, ByVal Y As Integer) As Boolean
  Dim L As Byte
  For L = 0 To MaxLayers
    If Terrain(Matrix(X, Y, L)).Name = "door" Then
      IsDoorAt = True
      Exit Property
    End If
  Next L
End Property

'clears array of text lines
Private Sub Class_Initialize()
  ReDim MyTextLine(0)
End Sub

'for testing only. shows some stats
Public Sub ShowDensityStatistics()
  Dim X As Integer
  Dim Y As Integer
  Dim L As Byte
  Dim TT As Long
  Dim C() As Long
  ReDim C(0 To MaxLayers)
  For X = 0 To MaxDim
    For Y = 0 To MaxDim
      If Matrix(X, Y, 0) > 0 Then
        If Not TerrainNameAt(X, Y) = DefaultTileName Then
          C(0) = C(0) + 1
        End If
      End If
      For L = 1 To MaxLayers
        If Matrix(X, Y, L) > 0 Then
          C(L) = C(L) + 1
        End If
      Next
    Next
  Next
  TT = CLng(MaxDim + 1) * CLng(MaxDim + 1)
  MsgBox "Layer 0 Density: " & C(0) / TT & vbCrLf _
       & "Layer 1 Density: " & C(1) / TT & vbCrLf _
       & "Layer 2 Density: " & C(2) / TT & vbCrLf _
       & "Layer 3 Density: " & C(3) / TT & vbCrLf _
       , vbOKOnly + vbInformation, "Density Statistics"
End Sub

'I use words because VB doesn't support unsigned INTs (dah!)

'returns Least Significant Byte
Private Function LSB(ByVal Word As Long) As Byte
  LSB = (Word And 255)
End Function

'returns Most Significant Byte
Private Function MSB(ByVal Word As Long) As Byte
  MSB = (Word And 65280) \ 256
End Function

'caps at max
Private Function OfMax(Value, Maximum)
  If Value > Maximum Then OfMax = Maximum Else OfMax = Value
End Function

'caps at min
Private Function OfMin(Value, Minimum)
  If Value < Minimum Then OfMin = Minimum Else OfMin = Value
End Function

'returns a number between A and B
Private Function Rand(ByVal A As Long, ByVal B As Long) As Long
  Rand = Rnd * (B - A) + A
End Function

'returns highest
Private Function Max(A, B)
  If A > B Then Max = A Else Max = B
End Function
