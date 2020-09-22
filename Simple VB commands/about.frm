VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About :: Simple VB commands :: Arzynik.com"
   ClientHeight    =   3345
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2308.779
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4215
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://www.arzynik.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   135
      Picture         =   "about.frx":058A
      Top             =   720
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   1560
      Picture         =   "about.frx":1165
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VB@arzynik.com"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ipliments many simple VB commands in one app."
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Label lblComments 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ceated by Devin Smith. To be used for educational purposes only."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The about box is pretty self explanitory

'________Begin about________

Private Class As New Class1

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Private Sub Label1_Click()
    Class.LaunchWebBrowser Me.hwnd, "http://www.arzynik.com"
End Sub

