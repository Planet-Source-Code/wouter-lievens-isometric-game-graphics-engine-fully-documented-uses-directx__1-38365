Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const KEY_DOWN As Integer = 4096
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -27
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Type PointAPI
  X As Long
  Y As Long
End Type

Public Map As New clsMapEngine
Public FPS As Long
Public LastFPS As Long
Public SelectedTile As Byte

Public AppPath As String

Public Mouse As PointAPI

Public DrawRange As Byte

Public Function Collision(ByVal X As Long, ByVal Y As Long _
                        , ByVal X1 As Long, ByVal Y1 As Long _
                        , ByVal X2 As Long, ByVal Y2 As Long) As Boolean
  'checks if X,Y is inside the specified rectangle
  Dim Temp As Long
  If X1 > X2 Then
    Temp = X2
    X2 = X1
    X1 = Temp
  End If
  If Y1 > Y2 Then
    Temp = Y2
    Y2 = Y1
    Y1 = Temp
  End If
  Collision = ((X >= X1 _
            And X <= X2 _
            And Y >= Y1 _
            And Y <= Y2 _
            ) = True)
End Function

Public Sub InitMap()
  Set Map.DestObj = frmMain.Board
  Map.InitDirectX False
  Map.Init
  Map.TileWidth = 48
  Map.MaxDim = 199
  Map.MaxLayers = 3
  Map.RoofLayer = 3
  Map.LoadTerrain AppPath & "\data\terrain.dat", AppPath & "\gfx\tiles\"
  Map.DefaultTileName = "grass"
  Map.FillByName "grass"
  Map.VisionX = 26
  Map.VisionY = 26
  Map.FocusX = 0
  Map.FocusY = 0
End Sub

Public Sub Main()
  AppPath = App.Path
  
  Randomize
  
  InitMap
  
  frmControls.FillNameList
  frmMain.WindowState = vbMaximized
  frmMain.Show
  frmControls.Show
  DoEvents
  
  If Command <> "" Then Map.LoadMap Command
End Sub

Public Sub RunGame()
  CheckKeys
  
  Map.ClearObjects
  
  AddSelector
  
  Map.AddObject AppPath & "\gfx\misc\focus.bmp", Map.FocusX, Map.FocusY, True
  
  Map.DrawMap
  
  frmControls.StatsLbl.Caption = LastFPS & "   [" & Mouse.X & ", " & Mouse.Y & "]   ObjectCount: " & Map.ObjectCount
  FPS = FPS + 1
End Sub

Public Sub CheckKeys()
  Dim CurM As PointAPI
  Dim K As String
  
  Dim X As Integer
  Dim Y As Integer
  
  GetCursorPos CurM
  ScreenToClient frmMain.hwnd, CurM
  
  If GetKeyState(vbKeyF1) And KEY_DOWN Then
    DrawRange = 0
  End If
  If GetKeyState(vbKeyF2) And KEY_DOWN Then
    DrawRange = 1
  End If
  If GetKeyState(vbKeyF3) And KEY_DOWN Then
    DrawRange = 2
  End If
  If GetKeyState(vbKeyF4) And KEY_DOWN Then
    DrawRange = 3
  End If
  
  If frmControls.LayerScroll.ListIndex = -1 Then frmControls.LayerScroll.ListIndex = 0
  
  If GetForegroundWindow = frmControls.hwnd Then Exit Sub
  
  If Collision(CurM.X, CurM.Y, 0, 0, frmMain.Board.Width, frmMain.Board.Height) Then
    Mouse.X = Map.GetTileX(CurM.X, CurM.Y)
    Mouse.Y = Map.GetTileY(CurM.X, CurM.Y)
    
    If GetKeyState(vbLeftButton) And KEY_DOWN Then
      For X = -DrawRange To DrawRange
        For Y = -DrawRange To DrawRange
          Map.Edit Mouse.X + X, Mouse.Y + Y, frmControls.LayerScroll.ListIndex, SelectedTile
        Next Y
      Next
    End If
    If GetKeyState(vbRightButton) And KEY_DOWN Then
      K = InputBox("Enter Key value", "Key", Map.KeyAt(Mouse.X, Mouse.Y))
      K = Val(K)
      Map.KeyAt(Mouse.X, Mouse.Y) = K
    End If
  End If
  
  If GetKeyState(vbKeyLeft) And KEY_DOWN Then
    Map.ScrollWest
  End If
  If GetKeyState(vbKeyRight) And KEY_DOWN Then
    Map.ScrollEast
  End If
  If GetKeyState(vbKeyUp) And KEY_DOWN Then
    Map.ScrollNorth
  End If
  If GetKeyState(vbKeyDown) And KEY_DOWN Then
    Map.ScrollSouth
  End If
End Sub

Public Sub AddSelector()
  Dim X As Integer
  Dim Y As Integer
  For X = -DrawRange To DrawRange
    For Y = -DrawRange To DrawRange
      Map.AddObject AppPath & "\gfx\misc\selector.bmp", Mouse.X + X, Mouse.Y + Y, True
    Next
  Next
End Sub
