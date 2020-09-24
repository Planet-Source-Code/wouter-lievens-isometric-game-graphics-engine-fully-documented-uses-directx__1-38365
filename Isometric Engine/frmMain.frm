VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RPG Map Editor"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ForeColor       =   &H0000FFFF&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   743
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   Visible         =   0   'False
   Begin VB.Timer GameTimer 
      Interval        =   1
      Left            =   9840
      Top             =   7440
   End
   Begin VB.PictureBox Board 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   960
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Timer Animator 
      Interval        =   500
      Left            =   9360
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   7440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Animator_Timer()
  Map.AnimateTiles 'drives the animator of the engine
End Sub

Private Sub Form_Load()
  Board.Left = 0
  Board.Top = 0
  Board.Width = frmMain.ScaleWidth
  Board.Height = frmMain.ScaleHeight
End Sub

Private Sub Form_Resize()
  Me.WindowState = vbMaximized 'keep form maximized
End Sub

Private Sub GameTimer_Timer()
  RunGame 'drives the game loop
End Sub

Private Sub Timer1_Timer()
  'counts the frames per second
  LastFPS = FPS / (Timer1.Interval * 0.001)
  FPS = 0
End Sub
