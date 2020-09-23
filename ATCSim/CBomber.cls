VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for CBomb

Private imgBomb As Image

Public Event BombPositionChanged(X As Integer, Y As Integer, W As Integer, H As Integer)

Property Let imagebomb(bomb As Image)
    Set imgBomb = bomb
End Property

Public Sub GoDown(plane As CPlane, ByVal FormHeight As Integer)
 Dim j As Integer
 imgBomb.Left = plane.X + 0.5 * plane.W
 imgBomb.Top = plane.Y + plane.H
 imgBomb.Visible = True
 Do While imgBomb.Top < FormHeight
    imgBomb.Top = imgBomb.Top + 5
    RaiseEvent BombPositionChanged(imgBomb.Left, imgBomb.Top, imgBomb.Width, imgBomb.Height)
    'pause
    For j = 1 To 2000
    Next j
 Loop
 imgBomb.Visible = False
End Sub

Public Sub destroy()
 imgBomb.Visible = False
End Sub

Public Sub Class_Terminate()
 Set imgBomb = Nothing
End Sub

