Attribute VB_Name = "MMain"
Option Explicit

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Integer, ByVal Y As Integer) As Long

Public OnProgress As Boolean

Public Function HiColor(Color As Long, WhiteFactor As Byte) As Long
Dim Red As Long
Dim Green As Long
Dim Blue As Long

  Blue = (Color And RGB(0, 0, 255)) + RGB(0, 0, WhiteFactor)
  If Blue > RGB(0, 0, 255) Then Blue = RGB(0, 0, 255)
  
  Green = (Color And RGB(0, 255, 0)) + RGB(0, WhiteFactor, 0)
  If Green > RGB(0, 255, 0) Then Green = RGB(0, 255, 0)
  
  Red = (Color And RGB(255, 0, 0)) + RGB(WhiteFactor, 0, 0)
  If Red > RGB(255, 0, 0) Then Red = RGB(255, 0, 0)
  
  HiColor = Red + Green + Blue
End Function

Public Function DarkColor(Color As Long) As Long
Dim Red As Long
Dim Green As Long
Dim Blue As Long

  Blue = (Color And RGB(0, 0, 255)) \ 2 And RGB(0, 0, 255)
  Green = (Color And RGB(0, 255, 0)) \ 2 And RGB(0, 255, 0)
  Red = (Color And RGB(255, 0, 0)) \ 2 And RGB(255, 0, 0)
  DarkColor = Red + Green + Blue
End Function


