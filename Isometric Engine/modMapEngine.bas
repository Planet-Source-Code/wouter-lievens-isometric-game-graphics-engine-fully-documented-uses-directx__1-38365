Attribute VB_Name = "modMapEngine"
Option Explicit

Public Enum DirectionList
  dirNorth = 0
  dirNorthEast = 1
  dirEast = 2
  dirSouthEast = 3
  dirSouth = 4
  dirSouthWest = 5
  dirWest = 6
  dirNorthWest = 7
  dirNowhere = 255
End Enum

Public Type Vector2D
  X As Integer
  Y As Integer
End Type

Public Function ToVector(ByVal X As Integer, ByVal Y As Integer) As Vector2D
  ToVector.X = X
  ToVector.Y = Y
End Function

