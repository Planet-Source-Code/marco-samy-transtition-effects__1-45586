VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text FXShows Engine"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   652
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Progress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2400
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2400
      List            =   "Form1.frx":0049
      TabIndex        =   4
      Text            =   "[None] = 1"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   5280
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      Picture         =   "Form1.frx":0228
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5040
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   2520
      Picture         =   "Form1.frx":22D3
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "progress"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "EffectExec.bas         WinDIctl.bas             FXShow.cls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "This Engine Only Requires The 3 files :"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "I also programmed the module (WinDIctl) but from long time befor 17 / 5 / 2003"
      Height          =   615
      Left            =   7560
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Full Programmed By Marco Samy - Started 17 / 5 / 2003 and Ended 19 / 5 / 2003, and still under update."
      Height          =   855
      Left            =   7560
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":3A91
      Height          =   1335
      Left            =   7560
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Choose FX"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iFFx As FXShows
Dim cIdx As Long
Private Sub Command1_Click()
Set iFFx = New FXShows
'how to use ower engine?

'first set the first picture box
Set iFFx.SourceDC1 = Picture1
'then set the second picture box
Set iFFx.SourceDC2 = Picture2
'set the distination  directory which final photo will be drawn on
Set iFFx.DestDC1 = Picture3
'select FX effect
iFFx.CurrectEffect = Combo1.ListIndex + 1
'must write this befor begin a new effect
iFFx.PrepareForNew

'our precentage value is stored in this varible
cIdx = 0

'begin
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0 'effect none
End Sub

Private Sub Timer1_Timer()
If cIdx >= 100 Then Timer1.Enabled = False: Exit Sub 'stop when reach 100 (%)
cIdx = cIdx + 2.5 'current percentage
'WHEN YOU USE LARGE PHOTO MAKE THE ADDITION LARGE TOO
'TO MAKE THE SHOW MORE FLIXABLE
'FOR EXAMPLE, IF USE LARGE PHOTO, REPLACE THE PREVIOUS LINE WITH:
'cIdx = cIdx + 5

Progress.Cls 'clear progress box

'draw current progress
Progress.Line (0, 0)-(Progress.ScaleWidth * cIdx / 100, Progress.ScaleHeight), Progress.FillColor, BF

'this is the function for drawing the effect , just type it and then type required FX effect index
iFFx.ExecuteEffect cIdx
'check to view photo
Picture3.Refresh
End Sub
