VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyPod"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox Filmain 
      Height          =   285
      Left            =   120
      Pattern         =   "*.wav;*.mp3;*.wma;*.wax;*.mid;*.midi;*.rmi;*.au;*.snd;*.aif;*.aifc;*.aiff"
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.DirListBox Dirmain 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.DriveListBox Drvmain 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin IPod.ThumbWheel THWmain 
      Height          =   1800
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      Size            =   -1  'True
   End
   Begin IPod.LCD LCDmain 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      Begin IPod.Menu Mnumain 
         Height          =   1620
         Left            =   30
         TabIndex        =   8
         Top             =   390
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   2858
      End
      Begin IPod.StatusBar barmain 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   238
         Max             =   359
      End
      Begin IPod.BatteryLevel BatMain 
         Height          =   150
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   265
      End
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   1
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   4200
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   2
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      Top             =   4200
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   3
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrDir As String
Const iPod_Green As Long = &HC8DDC1
Const MobilePhile_Blue As Long = 13514752

Private Sub btnmain_MouseShortClick(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Index
        Case 1
            Dim temp As String
            temp = CurrDir
            CurrDir = GetMenu(CurrDir, "..\")
            If StrComp(CurrDir, temp, vbTextCompare) <> 0 Then MainMenu LCDmain, Mnumain, CurrDir
        Case 2: MediaPlay
    End Select
End Sub

Private Sub btnmain_MouseStillDown(Index As Integer, x As Single, Y As Single)
    btnmain_MouseShortClick Index, 0, 0, x, Y
End Sub

Private Sub Form_Load()
    Set DrvBox = Me.Drvmain
    Set DirBox = Me.Dirmain
    Set FilBox = Me.Filmain

    MediaContainersHwnd Me.hwnd
    Alias = "MyPod"
    MainMenu LCDmain, Mnumain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MediaClose
End Sub

Private Sub THWmain_PodChangeClockWise(angle As Long)
    barmain.Value = barmain.Value + angle
    'scrmain.Value = barmain.Value / 2
    BatMain.Percent = barmain.Percent
    Mnumain.SelectedItem = Mnumain.SelectedItem + 1
End Sub

Private Sub THWmain_PodChangeCounterClockWise(angle As Long)
    barmain.Value = barmain.Value - angle
    'scrmain.Value = barmain.Value / 2
    BatMain.Percent = barmain.Percent
    Mnumain.SelectedItem = Mnumain.SelectedItem - 1
End Sub

Private Sub THWmain_ThumbClick()
    Dim temp As String, go As Boolean
    temp = Mnumain.GetItem(Mnumain.SelectedItem, True)
    go = Mnumain.GetItem(Mnumain.SelectedItem, False) = ">"
    If go Then
        CurrDir = GetMenu(CurrDir, temp)
        MainMenu LCDmain, Mnumain, CurrDir
    Else
        Execute CurrDir, temp
    End If
End Sub
