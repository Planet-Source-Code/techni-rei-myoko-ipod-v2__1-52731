VERSION 5.00
Begin VB.UserControl StatusBar 
   BackColor       =   &H00C8DDC1&
   CanGetFocus     =   0   'False
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   ScaleHeight     =   135
   ScaleWidth      =   2025
   ToolboxBitmap   =   "StatusBar.ctx":0000
   Begin VB.Line Linmain 
      Index           =   1
      X1              =   15
      X2              =   1905
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Linmain 
      Index           =   0
      X1              =   15
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image imgmain 
      Height          =   135
      Index           =   1
      Left            =   2000
      Picture         =   "StatusBar.ctx":0312
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgmain 
      Height          =   135
      Index           =   0
      Left            =   0
      Picture         =   "StatusBar.ctx":0354
      Top             =   0
      Width           =   30
   End
   Begin VB.Shape Shpmain 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   105
      Left            =   15
      Top             =   15
      Width           =   750
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private v_max As Long, v_value As Long
Public Event Change()

Public Property Let Percent(temp As Long)
    If temp >= 0 And temp <= Max Then value = temp * 0.01 * v_max
End Property
Public Property Let Max(temp As Long)
    If temp < 0 Then temp = 1
    v_max = temp
    UserControl_Resize
    RaiseEvent Change
End Property
Public Property Get Percent() As Long
    If v_max > 0 Then Percent = v_value / v_max * 100
End Property
Public Property Get Max() As Long
    Max = v_max
End Property
Public Property Let value(temp As Long)
    If temp > Max Then temp = Max
    If temp < 0 Then temp = 0
    v_value = temp
    UserControl_Resize
    RaiseEvent Change
End Property

Public Property Get value() As Long
    value = v_value
End Property
Public Property Let BackColor(temp As OLE_COLOR)
    UserControl.BackColor = temp
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Max = PropBag.ReadProperty("Max", 60)
    value = PropBag.ReadProperty("Value", 10)
    BackColor = PropBag.ReadProperty("BackColor", &HC8DDC1)
End Sub

Private Sub UserControl_Resize()
    Const minwidth As Long = 180
    UserControl.Height = 135
    If UserControl.Width < minwidth Then UserControl.Width = minwidth
    If UserControl.Width >= minwidth And Max > 0 Then
        imgmain(1).Left = UserControl.Width - imgmain(1).Width
        Linmain(0).x2 = imgmain(1).Left
        Linmain(1).x2 = Linmain(0).x2
        shpmain.Width = imgmain(1).Left * (value / Max)
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Max", v_max, 60
    PropBag.WriteProperty "Value", v_value, 10
    PropBag.WriteProperty "BackColor", UserControl.BackColor, &HC8DDC1
End Sub
