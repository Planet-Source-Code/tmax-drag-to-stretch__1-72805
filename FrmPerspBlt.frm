VERSION 5.00
Begin VB.Form FrmPerspBlt 
   BackColor       =   &H00000000&
   Caption         =   "Stretching ~"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdHelp 
      Caption         =   "Help"
      Height          =   615
      Left            =   13200
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdVisible 
      Caption         =   "Visible"
      Height          =   615
      Left            =   11520
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4860
      Left            =   11520
      Picture         =   "FrmPerspBlt.frx":0000
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3660
   End
   Begin VB.PictureBox PicShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   120
      ScaleHeight     =   495
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   751
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   5760
         MousePointer    =   15  'Size All
         TabIndex        =   9
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   8760
         MousePointer    =   15  'Size All
         TabIndex        =   8
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   8760
         MousePointer    =   15  'Size All
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5760
         MousePointer    =   15  'Size All
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   600
         MousePointer    =   15  'Size All
         TabIndex        =   5
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2160
         MousePointer    =   15  'Size All
         TabIndex        =   4
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2160
         MousePointer    =   15  'Size All
         TabIndex        =   3
         Top             =   720
         Width           =   255
      End
      Begin VB.Label LBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         MousePointer    =   15  'Size All
         TabIndex        =   2
         Top             =   720
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmPerspBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
' PerspBlt By Tmax
' Perspective Stretching
' Drag the Lbox's and stretching the images.
'*******************************************
Option Explicit
Dim StartX, StartY As Long
Dim StartHeight, EndHeight As Long
Dim StartWidth, EndWidth As Long
Dim OutWidth, OutHeight As Long
Dim OutXOffset, OutYOffset As Long
Dim Dx, Dy As Long
Dim ShowInfo As Boolean
Dim OnOff As Boolean

Private Sub CmdHelp_Click()
  If ShowInfo Then
      PicShow.Picture = LoadPicture(App.Path & "\stretching_info.jpg")
  Else
      PicShow.Picture = LoadPicture()
      DrawPic
  End If
  ShowInfo = Not ShowInfo
End Sub

Private Sub CmdVisible_Click()
Dim i%
For i% = 0 To 7
  LBox(i%).Visible = OnOff
Next
OnOff = Not OnOff
End Sub

Private Sub Form_Load()
ShowInfo = True
OnOff = False
DrawPic
End Sub

Private Sub LBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    PicShow.ScaleMode = 1
    Dx = x
    Dy = y
    PicShow.ScaleMode = 3
End If
End Sub

Private Sub LBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    PicShow.ScaleMode = 1
    LBox(Index).Left = LBox(Index).Left - (Dx - x)
    LBox(Index).Top = LBox(Index).Top - (Dy - y)
    PicShow.ScaleMode = 3
    UpDateLBox Index
End If
End Sub

Private Sub LBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dx = 0
Dy = 0
End Sub

Sub UpDateLBox(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0:
        LBox(3).Left = LBox(0).Left
    Case 1:
        LBox(2).Left = LBox(1).Left
    Case 2:
        LBox(1).Left = LBox(2).Left
    Case 3:
        LBox(0).Left = LBox(3).Left
    Case 4:
        LBox(5).Top = LBox(4).Top
    Case 5:
        LBox(4).Top = LBox(5).Top
    Case 6:
        LBox(7).Top = LBox(6).Top
    Case 7:
        LBox(6).Top = LBox(7).Top
    End Select
    DrawPic
End Sub

Sub DrawPic()
PicShow.Cls
' set up data for PerspBltX
StartX = LBox(0).Left + LBox(0).Width / 2
StartY = LBox(0).Top + LBox(0).Width / 2
StartHeight = LBox(3).Top - LBox(0).Top
EndHeight = LBox(2).Top - LBox(1).Top
OutWidth = LBox(1).Left - LBox(0).Left
OutYOffset = LBox(1).Top - LBox(0).Top
Call PerspBltX(PicShow.hdc, StartX, StartY, OutWidth, StartHeight, EndHeight, OutYOffset, Pic.hdc, Pic.ScaleWidth, Pic.ScaleHeight)

' set up data for PerspBltY
StartX = LBox(4).Left + LBox(4).Width / 2
StartY = LBox(4).Top + LBox(4).Width / 2
StartWidth = LBox(5).Left - LBox(4).Left
EndWidth = LBox(6).Left - LBox(7).Left
OutHeight = LBox(7).Top - LBox(4).Top
OutXOffset = LBox(7).Left - LBox(4).Left
Call PerspBltY(PicShow.hdc, StartX, StartY, StartWidth, EndWidth, OutHeight, OutXOffset, Pic.hdc, Pic.ScaleWidth, Pic.ScaleHeight)
End Sub

