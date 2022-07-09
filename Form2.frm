VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00511A13&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLegend 
      BackColor       =   &H00511A13&
      Caption         =   "Key to items shown by default?"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00511A13&
      ForeColor       =   &H00D8AEAB&
      Height          =   2055
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Text            =   "Form2.frx":15F942
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton btnQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A4512D&
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   340
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox chkStats 
      BackColor       =   &H00511A13&
      Caption         =   "Statistics shown by default?"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00511A13&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1440
      TabIndex        =   15
      Top             =   720
      Width           =   1575
      Begin VB.Label Label10 
         BackColor       =   &H00511A13&
         Caption         =   "S - Show/Hide stats"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00511A13&
         Caption         =   "Q - Quit"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00511A13&
         Caption         =   "N - Shoot"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00511A13&
         Caption         =   "M - Tractor Beam"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00511A13&
         Caption         =   "Space - Thrust"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00511A13&
         Caption         =   "X - Rotate right"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00511A13&
         Caption         =   "Z - Rotate left"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00A4512D&
         Caption         =   "Controls"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.OptionButton optRes5 
      BackColor       =   &H00511A13&
      Caption         =   "1600 x 1200"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optRes4 
      BackColor       =   &H00511A13&
      Caption         =   "1280 x 1024"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optRes3 
      BackColor       =   &H00511A13&
      Caption         =   "1024 x 768"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optRes2 
      BackColor       =   &H00511A13&
      Caption         =   "800 x 600"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optRes1 
      BackColor       =   &H00511A13&
      Caption         =   "640 x 480"
      ForeColor       =   &H00BA996D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnLaunch 
      Appearance      =   0  'Flat
      BackColor       =   &H00A4512D&
      Caption         =   "Launch"
      Height          =   340
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00511A13&
      BorderStyle     =   0  'None
      ForeColor       =   &H00BA996D&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
      Begin VB.OptionButton optBit2 
         BackColor       =   &H00511A13&
         Caption         =   "32 Bit"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optBit1 
         BackColor       =   &H00511A13&
         Caption         =   "16 Bit"
         ForeColor       =   &H00BA996D&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00A4512D&
         Caption         =   "Colour depth"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox txtYRes 
      Height          =   375
      Left            =   -240
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "1200"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtXRes 
      Height          =   375
      Left            =   -240
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "1600"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtColourDepth 
      Height          =   285
      Left            =   -240
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "16"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   4680
      Picture         =   "Form2.frx":1602C3
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00A4512D&
      Caption         =   "Resolution"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLaunch_Click()
    Form1.Show
End Sub

Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    optBit2 = True
    optRes5 = True
    chkStats = 1
    chkLegend = 1
End Sub



Private Sub optBit1_Click()
    txtColourDepth.Text = 16
End Sub

Private Sub optBit2_Click()
    txtColourDepth.Text = 32
End Sub

Private Sub optRes1_Click()
    txtXRes.Text = 640
    txtYRes.Text = 480
End Sub

Private Sub optRes2_Click()
    txtXRes.Text = 800
    txtYRes.Text = 600
End Sub

Private Sub optRes3_Click()
    txtXRes.Text = 1024
    txtYRes.Text = 768
End Sub

Private Sub optRes4_Click()
    txtXRes.Text = 1280
    txtYRes.Text = 1024
End Sub

Private Sub optRes5_Click()
    txtXRes.Text = 1600
    txtYRes.Text = 1200
End Sub

Private Sub Text1_Click()
    btnLaunch.SetFocus
End Sub
