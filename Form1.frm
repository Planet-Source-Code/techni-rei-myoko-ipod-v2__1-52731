VERSION 5.00
Begin VB.Form frmremote 
   BorderStyle     =   0  'None
   Caption         =   "Remote"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   45
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgmain 
      Height          =   375
      Index           =   3
      Left            =   0
      Tag             =   "2"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image imgmain 
      Height          =   375
      Index           =   2
      Left            =   0
      Tag             =   "2"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgmain 
      Height          =   375
      Index           =   1
      Left            =   360
      Tag             =   "3"
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgmain 
      Height          =   375
      Index           =   0
      Left            =   0
      Tag             =   "0"
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmremote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If Me.Picture <> 0 Then Call SetAutoRgn(Me)
setAlwaysOnTop Me.hwnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragform Me.hwnd
MoveForm Me
End Sub

Private Sub imgmain_Click(Index As Integer)
    If Index < 3 Then
        frmmain.btnmain_MouseClick imgmain(Index).tag
    Else
        Me.Visible = False
    End If
End Sub
