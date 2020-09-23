VERSION 5.00
Begin VB.UserControl iPodMenu 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   DataBindingBehavior=   1  'vbSimpleBound
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   ToolboxBitmap   =   "Menu.ctx":0000
   Begin IPod.LCD LCDmain 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      Border          =   0   'False
      Begin IPod.ScrollBar scrmain 
         Height          =   1815
         Left            =   2280
         Top             =   0
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   3201
         Max             =   4
         Value           =   4
         LargeChange     =   1
      End
   End
End
Attribute VB_Name = "iPodMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ItemHeight As Long = 18, WipeSpeed As Long = 10
Private Type MenuItem
     Lside As String
     Rside As String
     
     Underline As Boolean
End Type
Private MenuCount As Long, MenuList() As MenuItem, dir As Boolean, inter As Long, hide As Boolean
Private SelItem As Long, start As Long, onScreen As Long, locke As Boolean
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Sub RemoveItem(Index As Long)
    Dim temp As Long
    If MenuCount > 0 Then 'trying to cut down on exit sub/function
        If MenuCount = 1 Then
            ClearItems 'only one, just clear them all
        Else
            If Index < MenuCount - 1 Then 'must shift all from index onwards down one
                For temp = Index To MenuCount - 2
                    MenuList(temp) = MenuList(temp + 1)
                Next
            End If
            ReDim Preserve MenuList(MenuCount - 1) 'delete last menuitem
            MenuCount = MenuCount - 1
            If selecteditem >= Index Then selecteditem = selecteditem - 1 'change selecteditem if needed
        End If
        DrawMenu
    End If
End Sub
Public Property Get HideSelected() As Boolean
    HideSelected = hide
End Property
Public Property Let HideSelected(temp As Boolean)
    hide = temp
    DrawMenu
End Property
Public Function itemcount() As Long
    itemcount = MenuCount
End Function
Private Sub LCDmain_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub LCDmain_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub LCDmain_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Initialize()
    LCDmain.DoubleBuffer = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Public Function hwnd() As Long
    hwnd = UserControl.hwnd
End Function
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub
Private Sub LCDmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub
Public Property Let Interval(temp As Long)
    inter = temp
End Property
Public Property Get Interval() As Long
    Interval = inter
End Property
Public Property Let Locked(temp As Boolean)
    locke = temp
    If temp = False Then DrawMenu
End Property
Public Property Get Locked() As Boolean
    Locked = locke
End Property
Public Property Let Direction(temp As Boolean)
    dir = temp
End Property
Public Property Get Direction() As Boolean
    Direction = dir
End Property
Public Property Let BackColor(temp As OLE_COLOR)
    LCDmain.BackColor = temp
    UserControl.BackColor = temp
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = LCDmain.BackColor
End Property
Public Function GetItem(Index As Long, side As Boolean) As String
    If Index >= 0 And Index < MenuCount Then GetItem = IIf(side, MenuList(Index).Lside, MenuList(Index).Rside)
End Function
Public Function NewItem(ByVal text As String, Optional side As String) As Long
    NewItem = MenuCount
    If InStr(text, vbNewLine) > 0 Then text = Replace(text, vbNewLine, Empty)
    With scrmain
        .Max = MenuCount
        .LargeChange = onScreen
    End With
    MenuCount = MenuCount + 1
    ReDim Preserve MenuList(MenuCount)
    SetItem MenuCount - 1, text, side
End Function
Public Sub SetItem(Index As Long, text As String, Optional side As String)
    If Index >= 0 And Index < MenuCount Then
        With MenuList(Index)
            .Lside = text
            .Rside = side
        End With
    End If
    If Index >= start And Index <= start + onScreen - 1 Then DrawMenu
End Sub
Public Function SetSelectedItem(text As String) As Boolean
    Dim temp As Long
    For temp = 0 To MenuCount - 1
        If StrComp(MenuList(temp).Lside, text, vbTextCompare) = 0 Then
            selecteditem = temp
            SetSelectedItem = True
            Exit For
        End If
    Next
End Function
Public Sub ClearItems(Optional DoWipe As Boolean = False, Optional DoUnWipe As Boolean)
    MenuCount = 0
    ReDim MenuList(0)
    With scrmain
        .Value = 0
        .Max = 0
        .LargeChange = 0
    End With
    SelItem = 0
    start = 0
    If DoWipe Then Wipe
    LCDmain.ClearText
    LCDmain.LCDRefresh
    If DoUnWipe Then UnWipe
    DoEvents
End Sub
Public Sub Wipe()
    If inter = 0 Then inter = WipeSpeed
    Dim temp As Long, wid As Long
    wid = UserControl.Width
    For temp = 0 To wid / 15 Step inter
        If dir Then 'left
            LCDmain.Left = LCDmain.Left - inter
            scrmain.Left = scrmain.Left - inter
        Else 'right
            LCDmain.Left = LCDmain.Left + inter
            scrmain.Left = scrmain.Left + inter
        End If
        DoEvents
    Next
End Sub
Public Sub Pacman()
    Dim wid As Long
    wid = UserControl.Width / 15
    If dir = True Then
        LCDmain.Left = wid
        scrmain.Left = wid + LCDmain.Width
    Else
        LCDmain.Left = 0 - LCDmain.Width - scrmain.Width
        scrmain.Left = 0 - scrmain.Width
    End If
    dir = Not dir
    UnWipe
    dir = Not dir
End Sub

Public Sub UnWipe()
    Dim temp As Long, wid As Long
    wid = UserControl.Width
    If inter = 0 Then inter = WipeSpeed
    For temp = 0 To wid / 15 Step inter
        If LCDmain.Left < 0 Then
            LCDmain.Left = LCDmain.Left + inter
            scrmain.Left = scrmain.Left + inter
        Else 'right
            LCDmain.Left = LCDmain.Left - inter
            scrmain.Left = scrmain.Left - inter
        End If
        DoEvents
    Next
    'MsgBox LCDmain.Left
    LCDmain.Left = 0
    scrmain.Move wid - scrmain.Width
End Sub
Public Property Let selecteditem(Index As Long)
    If MenuCount > 0 Then
    If Index < 0 Then Index = (MenuCount + Index) Mod MenuCount
    If Index >= MenuCount Then Index = Index Mod MenuCount
    If SelItem = Index Then Exit Property
    SelItem = Index
    scrmain.Value = Index
    If Index > start + onScreen - 1 Then
        Do Until Index <= start + onScreen - 1
            start = start + 1
        Loop
    End If
    
    If (HideSelected And start <> Index) Or (Not HideSelected) Then
        If Index < start Then start = Index
        DrawMenu
    End If
    End If
End Property
Public Property Get selecteditem() As Long
    selecteditem = SelItem
End Property

Private Sub UserControl_Resize()
    Dim hit As Long, wid As Long
    UserControl.Height = (((UserControl.Height / 15) \ ItemHeight) * ItemHeight) * 15
    
    hit = UserControl.Height
    wid = UserControl.Width
    
    LCDmain.Move 0, 0, wid, hit
    scrmain.Move wid - scrmain.Width, 0, scrmain.Width, hit
    
    onScreen = (hit / 15) \ ItemHeight
    With scrmain
        .Max = MenuCount
        .LargeChange = onScreen
    End With
    DrawMenu
End Sub

Public Sub DrawMenu()
    If locke Then Exit Sub
    Dim temp As Long, y As Long, wid As Long, templ As Long, tempr As Long, hit As Long, tempstr As String
    If MenuCount = 0 Then Exit Sub
    Const WhiteSpace As Long = 4
    wid = (UserControl.Width - scrmain.Width) / 15
    LCDmain.ClearText
    For temp = start To start + onScreen - 1
        If temp >= 0 And temp < MenuCount Then
            hit = StringHeight(MenuList(temp).Lside & MenuList(temp).Rside)
            y = (temp - start) * ItemHeight
            tempr = StringWidth(MenuList(temp).Rside) + WhiteSpace
            If temp = SelItem And Not hide Then
                templ = StringWidth(MenuList(temp).Lside) + WhiteSpace
                LCDmain.DrawSquare 0, y, wid, 4, vbBlack, True 'Top Bar
                LCDmain.DrawSquare 0, y + ItemHeight - 2, wid, 2, vbBlack, True 'Bottom bar
                LCDmain.DrawSquare 0, y + WhiteSpace, 4, 12, vbBlack, True 'Left side middle bar
                LCDmain.DrawSquare templ, y, wid - templ - tempr, ItemHeight, vbBlack, True    'Middle middle bar
                LCDmain.DrawSquare wid - WhiteSpace, y + WhiteSpace, WhiteSpace, ItemHeight - WhiteSpace, vbBlack, True
            End If
            tempstr = Truncate(MenuList(temp).Lside, wid - WhiteSpace - WhiteSpace)
            If MenuList(temp).Underline Then LCDmain.DrawLine WhiteSpace / 2, y + ItemHeight - 2, StringWidth(tempstr) + WhiteSpace, 1, IIf(temp = SelItem, LCDmain.BackColor, vbBlack)
            LCDmain.PrintText tempstr, WhiteSpace, y + 4, temp = SelItem And Not hide
            LCDmain.PrintText MenuList(temp).Rside, wid - tempr, y + 4, temp = SelItem And Not hide
        End If
    Next
    LCDmain.LCDRefresh
    DoEvents
End Sub
Public Property Get Underline(Index As Long) As Boolean
    Underline = MenuList(Index).Underline
End Property
Public Property Let Underline(Index As Long, temp As Boolean)
    MenuList(Index).Underline = temp
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackColor = PropBag.ReadProperty("BackColor", &HC8DDC1)
    Direction = PropBag.ReadProperty("Direction", False)
    Interval = PropBag.ReadProperty("Interval", WipeSpeed)
    Locked = PropBag.ReadProperty("Locked", False)
    HideSelected = PropBag.ReadProperty("HideSelected", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", LCDmain.BackColor, &HC8DDC1
    PropBag.WriteProperty "Direction", dir, False
    PropBag.WriteProperty "Interval", inter, WipeSpeed
    PropBag.WriteProperty "Locked", locke, False
    PropBag.WriteProperty "HideSelected", hide, False
End Sub
