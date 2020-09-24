VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmControls 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controls"
   ClientHeight    =   3480
   ClientLeft      =   240
   ClientTop       =   4290
   ClientWidth     =   5280
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox LayerList 
      Appearance      =   0  'Flat
      Height          =   930
      ItemData        =   "frmControls.frx":0000
      Left            =   1680
      List            =   "frmControls.frx":0010
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton NullBtn 
      Caption         =   "Null Tile"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox LayerScroll 
      Height          =   315
      ItemData        =   "frmControls.frx":003E
      Left            =   120
      List            =   "frmControls.frx":004E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton InsideBtn 
      Caption         =   "Inside View"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton RandomizeBtn 
      Caption         =   "Randomize Tiles"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton FixBtn 
      Caption         =   "Fix Beach/Road"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ListBox TileNameList 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox SpriteNameList 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   3480
      TabIndex        =   0
      Tag             =   "0"
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   4440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "map"
      DialogTitle     =   "Select a file..."
   End
   Begin VB.CommandButton DensityBtn 
      Caption         =   "Show Stats"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton LoadBtn 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton SaveBtn 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton QuitBtn 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label StatsLbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tekst"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   8
      Top             =   3120
      Width           =   465
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FillNameList()
  Dim tAr() As String
  Dim n As Byte
  TileNameList.Clear
  tAr = Map.TerrainNameList
  For n = 1 To UBound(tAr)
    TileNameList.AddItem tAr(n)
  Next
End Sub

Private Sub DensityBtn_Click()
  Map.ShowDensityStatistics
End Sub

Private Sub FixBtn_Click()
  Map.FixBeaches
  Map.FixRoads
End Sub

Private Sub Form_Load()
  'keeps window on top
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  Dim n As Byte
  For n = 0 To LayerList.ListCount - 1
    LayerList.Selected(n) = True
  Next
End Sub

Private Sub InsideBtn_Click()
  If Map.InsideView Then Map.InsideView = False Else Map.InsideView = True
  If Map.InsideView Then
    InsideBtn.Caption = "Outside View"
  Else
    InsideBtn.Caption = "Inside View"
  End If
End Sub

Private Sub LayerList_Click()
  Map.LayerHidden(LayerList.ListIndex) = Not LayerList.Selected(LayerList.ListIndex)
End Sub

Private Sub LoadBtn_Click()
  FileDialog.DialogTitle = "Select a file to load..."
  FileDialog.ShowOpen
  If FileDialog.FileName <> "" Then Map.LoadMap FileDialog.FileName
End Sub

Private Sub NullBtn_Click()
  SelectedTile = 0
End Sub

Private Sub QuitBtn_Click()
  Map.UnloadDirectX
  End
End Sub

Private Sub RandomizeBtn_Click()
  Map.RandomizeTiles
End Sub

Private Sub SaveBtn_Click()
  FileDialog.DialogTitle = "Select a file to save to..."
  FileDialog.ShowSave
  If FileDialog.FileName <> "" Then Map.SaveMap FileDialog.FileName
End Sub

Private Sub SpriteNameList_Click()
  SelectedTile = Map.GetTerrainIDBySprite(SpriteNameList.Text)
  frmTile.TilePic.Picture = LoadPicture(AppPath & "\gfx\tiles\" & Map.TileSpriteName(SelectedTile) & ".bmp")
  frmTile.Show
  Me.Show
End Sub

Private Sub TileNameList_Click()
  Dim tAr() As Byte
  Dim n As Byte
  tAr = Map.AssembleTerrainList(TileNameList.Text)
  SpriteNameList.Clear
  For n = 0 To UBound(tAr)
    SpriteNameList.AddItem Map.TileSpriteName(tAr(n))
  Next
  SpriteNameList.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = True
End Sub
