VERSION 5.00
Begin VB.UserControl ScrollBar 
   BackColor       =   &H00C8DDC1&
   CanGetFocus     =   0   'False
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   135
   ScaleHeight     =   109
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   9
   ToolboxBitmap   =   "ScrollBar.ctx":0000
   Begin VB.Shape Shpmain 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Index           =   1
      Left            =   30
      Top             =   30
      Width           =   75
   End
   Begin VB.Shape Shpmain 
      Height          =   1635
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim v_min As Long, v_max As Long, v_value As Long, v_large As Long, isdown As Boolean
Public Event Scroll()

Private Sub UserControl_Resize()
If range >= 0 And Min <= Max And value >= Min And value <= Max And LargeChange <= range And LargeChange >= 0 Then
Dim wid As Double, hit As Double, temp As Double
wid = UserControl.Width / 15
hit = UserControl.Height / 15
shpmain(0).Move 0, 0, wid, hit

With shpmain(1)
    .Left = 2
    .Width = wid - 4
    .Height = (v_large / range) * (hit - 4)
    
    If range = 0 Or v_large = 0 Then wid = 0 Else wid = range / v_large   'Pages
    hit = shpmain(0).Height - 4 'Max height
    If wid = 0 Then temp = hit Else temp = hit / wid 'Pixels per page
    If v_large = 0 Then .top = 2 Else .top = 2 + ((value - Min) / v_large) * temp
    
    hit = shpmain(0).Height
    If .top + .Height > hit - 2 Then
        .top = hit - .Height - 2
    End If
End With


End If
End Sub

Private Function range() As Long
    range = Max - Min + 1
End Function

Public Property Let Min(temp As Long)
    If temp < Max Then
        v_min = temp
        If v_value < temp Then v_value = temp
        UserControl_Resize
    End If
End Property

Public Property Get Min() As Long
    Min = v_min
End Property

Public Property Let BackColor(temp As OLE_COLOR)
    UserControl.BackColor = temp
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let Max(temp As Long)
    If temp >= Min Then
        v_max = temp
        If v_value > temp Then v_value = temp
        UserControl_Resize
    End If
End Property

Public Property Get Max() As Long
    Max = v_max
End Property
Public Property Let value(temp As Long)
    If temp < Min Then temp = Min
    If temp > Max Then temp = Max
    v_value = temp
    UserControl_Resize
    RaiseEvent Scroll
End Property

Public Property Get value() As Long
    value = v_value
End Property
Public Property Let LargeChange(temp As Long)
    If temp < 1 Then temp = 1
    If temp > range Then temp = range
    v_large = temp
    UserControl_Resize
End Property

Public Property Get LargeChange() As Long
    LargeChange = v_large
End Property

Private Sub UserControl_InitProperties()
    Min = 0
    Max = 16
    value = 0
    LargeChange = 4
    BackColor = &HC8DDC1
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Min = PropBag.ReadProperty("Min", 0)
    Max = PropBag.ReadProperty("Max", 16)
    value = PropBag.ReadProperty("Value", 0)
    LargeChange = PropBag.ReadProperty("LargeChange", 4)
    BackColor = PropBag.ReadProperty("Backcolor", &HC8DDC1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Min", v_min, 0
    PropBag.WriteProperty "Max", v_max, 16
    PropBag.WriteProperty "Value", v_value, 0
    PropBag.WriteProperty "LargeChange", v_large, 4
    PropBag.WriteProperty "Backcolor", UserControl.BackColor, &HC8DDC1
End Sub
