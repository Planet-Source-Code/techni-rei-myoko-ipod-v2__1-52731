VERSION 5.00
Begin VB.UserControl Button 
   BackStyle       =   0  'Transparent
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   MaskColor       =   &H000000FF&
   MaskPicture     =   "Button.ctx":0000
   Picture         =   "Button.ctx":1302
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   40
   ToolboxBitmap   =   "Button.ctx":2604
   Begin VB.Timer Timermain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   3
      Left            =   97
      Picture         =   "Button.ctx":2916
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   2
      Left            =   97
      Picture         =   "Button.ctx":296C
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   1
      Left            =   97
      Picture         =   "Button.ctx":29C4
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgmain 
      Height          =   120
      Index           =   0
      Left            =   97
      Picture         =   "Button.ctx":2A22
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private picindex As Long, picvisible As Long, wasdown As Boolean, cx As Single, cy As Single
Public Event MouseShortClick(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseStillDown(X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub imgmain_Click(index As Integer)
    RaiseEvent MouseClick
End Sub

Private Sub imgmain_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X + imgmain(index).Left, Y + imgmain(index).top
End Sub

Private Sub imgmain_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X + imgmain(index).Left, Y + imgmain(index).top
End Sub

Private Sub imgmain_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X + imgmain(index).Left, Y + imgmain(index).top
End Sub

Private Sub TimerMain_Timer()
    If Not wasdown Then
        wasdown = True
        RaiseEvent MouseStillDown(cx, cy)
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent MouseClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimerMain.Enabled = True
    wasdown = False
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If TimerMain.Enabled Then
        cx = X
        cy = Y
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimerMain.Enabled = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not wasdown Then RaiseEvent MouseShortClick(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 600
    UserControl.Height = UserControl.Width
End Sub
Public Property Let Image(temp As Long)
    Dim temp2 As Long
    If temp >= 0 And temp <= 3 Then picindex = temp
    For temp2 = 0 To 3
        imgmain(temp2).Visible = picvisible And temp = temp2
    Next
End Property
Public Property Let ImageVisible(temp As Boolean)
    picvisible = temp
    imgmain(picindex).Visible = temp
End Property
Public Property Let Interval(temp As Long)
    If temp > 0 And temp <= 10000 Then TimerMain.Interval = temp
End Property
Public Property Get Interval() As Long
    Interval = TimerMain.Interval
End Property
Public Property Get Image() As Long
    Image = picindex
End Property
Public Property Get ImageVisible() As Boolean
    ImageVisible = picvisible
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Image = PropBag.ReadProperty("Image", 0)
    ImageVisible = PropBag.ReadProperty("ImageVisible", True)
    Interval = PropBag.ReadProperty("Interval", 100)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Image", picindex, 0
    PropBag.WriteProperty "ImageVisible", ImageVisible, True
    PropBag.WriteProperty "Interval", TimerMain.Interval, 100
End Sub
