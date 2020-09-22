VERSION 5.00
Begin VB.Form music 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Music :: Simple VB commands :: Arzynik.com"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command31 
         Caption         =   "Stop"
         Height          =   225
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Pause"
         Height          =   225
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Play"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "in order to use the music section you must inclue windows the media player componet, msdxm.ocx"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'in order to use the music section you must inclue
'windows the media player componet, msdxm.ocx

'________Begin music________

Private Class As New Class1

Private Sub Form_Load()                                 'Defines the sub when loaded
    Command32.Visible = True
    Command31.Visible = True
    'media.FileName = "At the Drive In - Enfilade.mp3"
    'media.Play
End Sub

Private Sub Command33_Click()
    'Play control handler
    'media.Play
    Command33.Visible = False
    Command32.Visible = True
    Command31.Visible = True
End Sub

Private Sub Command32_Click()
    'Pause control handler
    'media.Pause
    Command32.Visible = False
    Command33.Visible = True
End Sub

Private Sub Command31_Click()
    'Stop control handler
    'media.Stop
    'media.CurrentPosition = 0
    Command33.Visible = True
    Command32.Visible = False
    Command31.Visible = False
End Sub
