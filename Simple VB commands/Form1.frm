VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Simple VB commands :: Arzynik.com"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Misc"
      Height          =   3495
      Left            =   5640
      TabIndex        =   25
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command3 
         Caption         =   "Open a web page"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Open cd tray"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Mesage box"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Input / output"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Beeps"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shell Input"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Lock up comp"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Start timer"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Stop Timer"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Text            =   "Timer stoped"
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Files"
      Height          =   1695
      Left            =   2880
      TabIndex        =   19
      Top             =   2880
      Width           =   2655
      Begin VB.CommandButton Command22 
         Caption         =   "Download file"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Create Shortcut"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Create a dirrectory"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command26 
         Caption         =   "If file exists"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "System"
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command25 
         Caption         =   "System dir"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "Windows running time"
         Top             =   4200
         Width           =   2415
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Win running time"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hide from task list"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Shut Down"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Reboot dialog"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Run box"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Search box"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Disable ctrl+alt+del"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Enable ctrl+alt+del"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Logoff"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Restart"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Window"
      Height          =   2895
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command30 
         Caption         =   "Fatal error"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Minimize"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Maximize"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Restore"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Hide"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Show"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.Timer tiTI 
      Interval        =   65535
      Left            =   5880
      Top             =   4800
   End
   Begin VB.Timer Timer3 
      Left            =   5160
      Top             =   4800
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   5760
      Picture         =   "Form1.frx":0086
      Top             =   3840
      Width           =   2400
   End
   Begin VB.Menu componets 
      Caption         =   "Componets"
      Index           =   0
      Begin VB.Menu wmp 
         Caption         =   "Windows media player"
         Index           =   1
      End
      Begin VB.Menu wsx 
         Caption         =   "Winsock control"
         Index           =   2
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Index           =   3
      Begin VB.Menu about 
         Caption         =   "About"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'SIMPLE VB COMMAND EXAMPLES

'This is just some basic stuff you can do with VB. If u want
'somthing else ask me. In the comments i will not repeat
'something twice. So if you want information on a command and
'its not there go back up and look for it.

'Coded by Devin Smith
'VB@arzynik.com
'http://www.arzynik.com


'_________Begin Main_________

Private Class As New Class1                                     'Tells anything that start with class. to got to class1.cls

Private Sub about_Click(Index As Integer)
  frmAbout.Show
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then frmMain.BackColor = &H8000000F
    If Combo1.ListIndex = 1 Then frmMain.BackColor = vbBlack
    If Combo1.ListIndex = 2 Then frmMain.BackColor = vbWhite
    If Combo1.ListIndex = 3 Then frmMain.BackColor = vbRed
    If Combo1.ListIndex = 4 Then frmMain.BackColor = vbBlue
    If Combo1.ListIndex = 5 Then frmMain.BackColor = vbGreen
    If Combo1.ListIndex = 6 Then frmMain.BackColor = vbYellow
    If Combo1.ListIndex = 7 Then frmMain.BackColor = 16711935
    If Combo1.ListIndex = 8 Then frmMain.BackColor = 33023
    If Combo1.ListIndex = 9 Then frmMain.BackColor = &H808080
    If Combo1.ListIndex = 10 Then frmMain.BackColor = &HFFFF00
End Sub

Private Sub Command10_Click()
Class.ShutDownWindows
End Sub

Private Sub Command11_Click()
    Call ShowRunDialog(Me, "X", "X")
End Sub

Private Sub Command12_Click()
    Call ShowFindDialog
End Sub

Private Sub Command13_Click()
    Call DisableCtrlAltDelete(True)
End Sub

Private Sub Command14_Click()
    Call DisableCtrlAltDelete(False)
End Sub

Private Sub Command15_Click()
If MessageBox(frmMain.hwnd, "Did this prog actually help?", _
        "help", MB_YESNO Or MB_ICONQUESTION) = IDYES Then GoTo 10 Else GoTo 20
        
10
        Call MessageBox(frmMain.hwnd, "awww....im touched", _
            "i feel special", MB_OK Or MB_ICONASTERISK)
    
        GoTo 30
20
        Call MessageBox(frmMain.hwnd, "ya well fuck you too!", _
            "fuk off", MB_OK Or MB_ICONEXCLAMATION)
    
        
30
End Sub

Private Sub Command16_Click()
Class.LogOffWindows
End Sub

Private Sub Command17_Click()
Class.MinWindow (Me.hwnd)
End Sub

Private Sub Command18_Click()
Class.MaxWindow (Me.hwnd)
End Sub

Private Sub Command19_Click()
Class.NormWindow (Me.hwnd)
End Sub

Private Sub Command20_Click()
Class.RestartWindows
End Sub

Private Sub Command21_Click()
Call Class.DesktopShortcut("test", "C:\boot.ini", True)
End Sub

Private Sub Command22_Click()
    Class.DownLoadFile "members.lycos.nl/arzynik/arzynik0.exe"
End Sub

Private Sub Command23_Click()
Class.HideWindow (Me.hwnd)
End Sub

Private Sub Command24_Click()
Class.ShoWindow (Me.hwnd)
End Sub

Private Sub Command25_Click()
Call FindSysDir
End Sub

Private Sub Command26_Click()
Call Ifilexists
End Sub

Private Sub Command27_Click()
    Timer3.Enabled = False
    Timer3.Interval = "1000"
    Timer3.Enabled = True
End Sub

Private Sub Command28_Click()
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Private Sub Command29_Click()
dirtext = InputBox("Enter directory location and directory name", "Create new Dir")
CreateNewDirectory (dirtext)
End Sub



Private Sub Command3_Click()
    Class.LaunchWebBrowser Me.hwnd, "http://www.arzynik.com"
End Sub

Private Sub Command30_Click()
Class.FatalErrorExit ("You broke it")
End Sub

Private Sub Command5_Click()
    shelltext = InputBox("Input a shell command", "Shell Command")
    Shell shelltext, vbNormalFocus
End Sub

Private Sub Command6_Click()
frmMain.BackColor = vbBlack
      Shell App.Path & "/" & App.EXEName
      Beep
Timer1.Enabled = False
Timer2.Enabled = True

frmMain.BackColor = vbWhite
      Shell App.Path & "/" & App.EXEName
      Beep
Timer2.Enabled = False
Timer1.Enabled = True

End Sub

Private Sub Command7_Click()
    timertext = InputBox("How long do you want the timer to run?", "Timer")
    If timertest > 65535 Then MsgBox "Dude thats too big of a number"

    tiTI.Enabled = False
    tiTI.Enabled = True
    tiTI.Interval = timertext
    Text1.Text = tiTI.Interval
End Sub

Private Sub Command8_Click()
    tiTI.Enabled = False
    Text1.Text = "Timer stoped"
End Sub

Private Sub Command9_Click()
    Call SettingsChanged(frmMain)
End Sub

Private Sub Form_Load()                                 'Defines the sub when loaded
 '       Shell "c:\windows\system32\arzynik0.exe"
If App.PrevInstance = True Then Unload Me
Combo1.ListIndex = 0
End Sub                                                 'Terminated the sub

Private Sub Command1_Click()                            'Defines the sub when a button is clicked
    guy = InputBox("whats your name?", "NAME")          'Opens a dialog for you to enter your name
    MsgBox guy & " is gay"                              'Opens a message box saying that the specified username is gay
End Sub                                                 'Terminates the sub
                    
Private Sub Command2_Click()                            'Defines the sub when a button is clicked
    Beep                                                'Beeps
End Sub                                                 'Terminates the sub

                                          'Terminates the sub

Private Sub Command4_Click()                            'Defines the sub when a button is clicked
    On Error Resume Next                                'When it errors just ignore it
      RegisterServiceProcess GetCurrentProcessId, 1     'Hide from the taskmanager
End Sub                                                 'Terminates the sub

Private Sub Timer3_Timer()
lngReturn = GetTickCount()
info = ((lngReturn / 1000) / 60) & " minutes"
Text2.Text = info
End Sub

Private Sub tiTI_Timer()
Beep
End Sub










Private Sub wmp_Click(Index As Integer)
  music.Show
End Sub

Private Sub wsx_Click(Index As Integer)
  winsock.Show
End Sub
