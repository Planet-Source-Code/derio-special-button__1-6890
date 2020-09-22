VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   1020
   ScaleWidth      =   2940
   ToolboxBitmap   =   "Button.ctx":0000
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum ButtonStyle
  bNone = 0
  bShow = 1
  bPush = 2
End Enum

Public Event Click()

Private vCaption As String

Private Point As POINTAPI
Private dX As Long
Private dy As Long
Private vOnMe As Boolean
Private vClick As Boolean


Public Property Get Caption() As String
  Caption = vCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  vCaption = vNewValue
  DrawButton bNone
  PropertyChanged "Caption"
End Property

Private Sub PrintCaption(Style As ButtonStyle)
  With UserControl
    If Style = bNone Then
      .FontBold = False
      .CurrentX = (.Width - .TextWidth(vCaption)) \ 2
      .CurrentY = (.Height - .TextHeight(vCaption)) \ 2
    Else
      .FontBold = True
      .CurrentX = (.Width - .TextWidth(vCaption)) \ 2
      .CurrentY = (.Height - .TextHeight(vCaption)) \ 2
      If Style = bPush Then
        .CurrentX = .CurrentX + 15
        .CurrentY = .CurrentY + 15
      End If
    End If
  End With
  UserControl.Print vCaption
End Sub

Private Sub UserControl_InitProperties()
  vCaption = Extender.Name
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    vClick = True
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpPoint As POINTAPI

  If OnProgress Then Exit Sub
  
  If X > 0 And X < Width And Y > 0 And Y < Height Then
    If vOnMe Then Exit Sub
    vOnMe = True
    GetCursorPos tmpPoint
    If Not vClick Then
      DrawButton bShow
    Else
      DrawButton bPush
    End If
    
    DoEvents
    Point.X = tmpPoint.X
    Point.Y = tmpPoint.Y
    dX = tmpPoint.X - X \ Screen.TwipsPerPixelX
    dy = tmpPoint.Y - Y \ Screen.TwipsPerPixelY
    If Not vClick Then
      CheckMouse bShow
    Else
      CheckMouse bPush
    End If
    vOnMe = False
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  DrawButton bNone
  vClick = False
  If X > 0 And X < Width And Y > 0 And Y < Width Then
    RaiseEvent Click
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Caption = PropBag.ReadProperty("Caption", vCaption)
  ForeColor = PropBag.ReadProperty("ForeColor", UserControl.ForeColor)
  Set Font = PropBag.ReadProperty("Font", UserControl.Font)
End Sub

Private Sub UserControl_Resize()
  If Width Mod Screen.TwipsPerPixelX <> 0 Then
    Width = Width \ Screen.TwipsPerPixelX * Screen.TwipsPerPixelX
  End If
  If Height Mod Screen.TwipsPerPixelY <> 0 Then
    Height = Height \ Screen.TwipsPerPixelY * Screen.TwipsPerPixelY
  End If
  DrawButton bNone
End Sub

Private Sub CheckMouse(OldStyle As ButtonStyle)
Dim MouseLeave As Boolean
Dim PosX As Long
Dim PosY As Long
Dim Button As Long
Dim Extra As Long

  If OnProgress Then Exit Sub
  OnProgress = True
  Do
    GetCursorPos Point
    With Point
      If Not (.X >= dX And .X <= dX + Width \ Screen.TwipsPerPixelX And _
              .Y >= dy And .Y <= dy + Height \ Screen.TwipsPerPixelY) Then
        If Not vClick Then
          If OldStyle <> bNone Then
            DrawButton bNone
            OldStyle = bNone
          End If
        
        Else
          DrawButton bShow
          OldStyle = bShow
        End If
        MouseLeave = True
      
      Else
        DoEvents
        If vClick Then
          If OldStyle <> bPush Then
            DrawButton bPush
            OldStyle = bPush
          End If
        
        Else
          If OldStyle <> bShow Then
            DrawButton bShow
            OldStyle = bShow
          End If
        End If
      End If
    End With
  Loop Until MouseLeave
  OnProgress = False
End Sub

Private Sub DrawButton(Style As ButtonStyle)
Dim I As Integer
Dim J As Integer
  
  Cls
  GetBackPicture
  PrintCaption Style
  Select Case Style
    Case bShow
      For J = 0 To Height \ Screen.TwipsPerPixelY - 1
        For I = 0 To Width \ Screen.TwipsPerPixelX
         SetPixel hDC, (I), (J), HiColor(GetPixel(hDC, (I), (J)), 30)
        Next
      Next J
     
      For I = 0 To Height \ Screen.TwipsPerPixelY
        SetPixel hDC, (0), (I), HiColor(GetPixel(hDC, (0), (I)), 50)
      Next
    
      For I = 0 To Width \ Screen.TwipsPerPixelX
        SetPixel hDC, (I), (Height \ Screen.TwipsPerPixelY - 1), DarkColor(GetPixel(hDC, (I), (Height \ Screen.TwipsPerPixelY - 1)))
      Next
    
      For I = 0 To Height \ Screen.TwipsPerPixelY
        SetPixel hDC, (Width \ Screen.TwipsPerPixelX - 1), (I), DarkColor(GetPixel(hDC, (Width \ Screen.TwipsPerPixelX - 1), (I)))
      Next
    
    Case bPush
      For I = 0 To Width \ Screen.TwipsPerPixelX
        SetPixel hDC, (I), (0), DarkColor(GetPixel(hDC, (I), (0)))
      Next
    
      For I = 0 To Height \ Screen.TwipsPerPixelY
        SetPixel hDC, (0), (I), DarkColor(GetPixel(hDC, (0), (I)))
      Next
      
      For J = 1 To Height \ Screen.TwipsPerPixelY - 1
        For I = 1 To Width \ Screen.TwipsPerPixelX
         SetPixel hDC, (I), (J), HiColor(GetPixel(hDC, (I), (J)), 30)
        Next
      Next J
     
      For I = 0 To Width \ Screen.TwipsPerPixelX
        SetPixel hDC, (I), (Height \ Screen.TwipsPerPixelY - 1), HiColor(GetPixel(hDC, (I), (Height \ Screen.TwipsPerPixelY - 1)), 50)
      Next
    
      For I = 0 To Height \ Screen.TwipsPerPixelY
        SetPixel hDC, (Width \ Screen.TwipsPerPixelX - 1), (I), HiColor(GetPixel(hDC, (Width \ Screen.TwipsPerPixelX - 1), (I)), 50)
      Next
    
  End Select
  PrintCaption Style
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", vCaption
  PropBag.WriteProperty "ForeColor", UserControl.ForeColor
  PropBag.WriteProperty "Font", UserControl.Font
End Sub

Private Sub GetBackPicture()
  On Error Resume Next
  PaintPicture Parent.Picture, _
               0, 0, Width, Height, _
               Extender.Left, Extender.Top, _
               Width, Height, opcode:=vbSrcCopy
  On Error GoTo 0
End Sub

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  UserControl.ForeColor = vNewValue
  DrawButton bNone
  PropertyChanged "ForeColor"
End Property

Public Property Get Font() As IFontDisp
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal vNewValue As IFontDisp)
  Set UserControl.Font = vNewValue
  DrawButton bNone
  PropertyChanged "Font"
End Property
