VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'THE MEAT OF THE PROGRAM

'Most of the code in this app is here in the class file.
'I happen to like class files better than modules so i
'used a class. So no bitching. And i didnt comment on all
'of it. Alot of it seemed a little ovious or redundunt. So
'have fun.

'________Begin Class1________

Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long 'Declaring stuff for hide from task
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function IsWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetParentAPI& Lib "user32" Alias "GetParent" (ByVal hwnd As Long)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Declaring stuff for the Find Dialog
Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long 'Declaring stuff for the Reboot Dialog
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long 'Declaring stuff for Disabling the CTRL+ALT+DEL box
Private Declare Function messagebox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long 'Declaring stuff for the Message Box
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal UnknownP1 As Long, ByVal UnknownP2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long 'Declaring stuff for the Run Dialog Box
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long

'Declaring stuff for the winsock control
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)

'Show Window Functions
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const SW_NORMAL = 1

Private Const SystemChangeRestart = 4

'Message Box Types
Private Const MB_ABORTRETRYIGNORE = &H2& 'Abort, Retry, Ignore
Private Const MB_YESNO = &H4& ' Yes and No
Private Const MB_YESNOCANCEL = &H3& 'Yes, No, Cancel
Private Const MB_RETRYCANCEL = &H5& 'Retry and Cancel
Private Const MB_OKCANCEL = &H1& 'Ok and Cancel
Private Const MB_OK = &H0& 'Just OK

'Icons
Private Const MB_ICONSTOP = &H10& 'Stop Icon
Private Const MB_ICONQUESTION = &H20& 'Question Mark Icon
Private Const MB_ICONASTERISK = &H40& 'Asterisk Icon
Private Const MB_ICONEXCLAMATION = &H30& 'Exclamation Icon

'Button Types
Private Const IDYES = 6 'Yes Button
Private Const IDNO = 7 'No Button
Private Const IDABORT = 3 'Abort Button
Private Const IDCANCEL = 2 'Cancel Button
Private Const IDIGNORE = 5 'Ignore Button
Private Const IDRETRY = 4 'Retry Button
Private Const IDOK = 1 'Ok Button

'Declaring stuff for Shutdown Windows
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Const EWX_FORCE = 4
Private Const EWX_LOGOFF = 0
Private Const EWX_REBOOT = 2
Private Const EWX_SHUTDOWN = 1

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
Private Type SECURITY_ATTRIBUTES
    nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Sub DownLoadFile(sUrl As String)
    Dim DL As Long
    If sUrl$ = "" Then sUrl$ = strUrl$
    If strUrl$ = "" Then strUrl$ = sUrl$
    If Left(strUrl$, 4) <> "http" Then strUrl$ = "http://" & strUrl$ 'See if  there's a http:// at the begining if not add it
        DL& = DoFileDownload(StrConv(strUrl$, vbUnicode))
End Sub
  
Public Sub RestartWindows()
    'reboot the computer
    Ret& = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
End Sub

Public Sub LogOffWindows()
    'LogOff  the computer
    Call ExitWindowsEx(EWX_LOGOFF, 0)
End Sub

Public Sub ShutDownWindows()
    'Shut Down  the computer
    Call ExitWindowsEx(EWX_FORCE Or EWX_SHUTDOWN, 0)
End Sub

Public Sub FatalErrorExit(ErrorMSG As String)
    'Creates a fatal error
    On Error GoTo errHandler
    '0 = white fatal error
    '1 =
    FatalAppExit 1, ErrorMSG$
    Exit Sub
errHandler:
    Debug.Print "Error Source:", Err.Source
    Debug.Print "Error Description:", Err.Description
    Debug.Print "Error Number:", Err.Number
End Sub

Sub MaxWindow(hwnd)
    'Maximize a Window
    Dim Max
    Max = ShowWindow(hwnd, SW_MAXIMIZE)
End Sub

Sub MinWindow(hwnd)
    'Minimizes a Window
    Dim min
    min = ShowWindow(hwnd, SW_MINIMIZE)
End Sub

Sub NormWindow(hwnd)
    'Minimizes a Window
    Dim nor
    nor = ShowWindow(hwnd, SW_NORMAL)
End Sub

Sub HideWindow(hwnd)
    'Hides a Window
    Dim hid
    hid = ShowWindow(hwnd, SW_HIDE)
End Sub

Sub ShoWindow(hwnd)
    'Shows a Window
    Dim sho
    sho = ShowWindow(hwnd, SW_SHOW)
End Sub

Public Sub LaunchWebBrowser(lhwnd As Long, Site As String)
    'Launch Web Browser To a site
    On Error Resume Next
    Dim IE As Long
    IE& = ShellExecute(lhwnd, "Open", Site, vbNullString, vbNullString, SW_MAXIMIZE)
End Sub

Public Sub ShowFindDialog(Optional InitialDirectory As String)
    'Code for the Find Dialog box
    ShellExecute 0, "find", _
    IIf(InitialDirectory = "", "", InitialDirectory), _
    vbNullString, vbNullString, SW_SHOW
End Sub

Public Sub SettingsChanged(FormName As Form)
    'Code for the Reboot Dialog box
    SHRestartSystemMB FormName.hwnd, vbNullString, SystemChangeRestart
End Sub

Sub DisableCtrlAltDelete(bDisabled As Boolean)
    'Code for Disabling the CTRL+ALT+DEL dialog box
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Public Sub ShowRunDialog(ByRef CallingForm As Form, Optional Title As String, Optional Description As String)
    'Code for the Run Dialog Box
    If Title = "" Then Title = "Run"
    If Description = "" Then Description = "Type the name of a program to open, " & "then click OK when finished."
    SHRunDialog CallingForm.hwnd, 0, 0, Title, Description, 0
End Sub

Public Sub FindSysDir()
    'Finds the windows system directory and opend it in explorer
    Dim WindirS As String * 255
    Dim SysDir
    SysDir = Left(WindirS, GetSystemDirectory(WindirS, 255))
    Shell "explorer " & SysDir
End Sub
Public Sub Ifilexists()
    'Determine if a file exists in the system directory
        Dim thedir
        Dim thefile
        thedir = "c:\"
        thefile = "autoexec.bat"
    If CreateObject("Scripting.FileSystemObject").FileExists(thedir & thefile) Then
        MsgBox thedir & thefile & " is there."
    Else
    MsgBox thedir & thefile & " is not there."
        'This is an easy way to copy and register an ocx file
        'FileSystem.FileCopy "Mscomctl.ocx", SysDir & "/Mscomctl.ocx"
        'Shell "regsvr32.exe Mscomctl.ocx"
    End If
End Sub

Public Sub CreateNewDirectory(NewDirectory As String)
    'Code that creates a new directory
    Dim sDirTest As String
    Dim SecAttrib As SECURITY_ATTRIBUTES
    Dim bSuccess As Boolean
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String
    iFlag = 0
    sPath = NewDirectory
    
    If Right(sPath, Len(sPath)) <> "\" Then
        sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left(sPath, iCounter)
        sDirTest = Dir(sTempDir)
        iCounter = iCounter + 1
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop
End Sub

Public Sub HideTask()
    On Error Resume Next
        RegisterServiceProcess GetCurrentProcessId, 1
End Sub

Public Sub Wintime()
    lngReturn = GetTickCount()
    info = ((lngReturn / 1000) / 60) & " minutes"
    frmMain.Text2.Text = info
End Sub

Public Sub CompLok()
    'Will lock up your computer
    frmMain.BackColor = vbBlack
        Shell App.Path & "/" & App.EXEName 'Opens itself
        Beep
    frmMain.Timer1.Enabled = False
    frmMain.Timer2.Enabled = True
    frmMain.BackColor = vbWhite
        Shell App.Path & "/" & App.EXEName
        Beep
    frmMain.Timer2.Enabled = False
    frmMain.Timer1.Enabled = True
End Sub

Public Sub bgChange()
    If frmMain.Combo1.ListIndex = 0 Then frmMain.BackColor = &H8000000F 'Changes the forms bgcolor depending on what is selected in the combo box
    If frmMain.Combo1.ListIndex = 1 Then frmMain.BackColor = vbBlack
    If frmMain.Combo1.ListIndex = 2 Then frmMain.BackColor = vbWhite
    If frmMain.Combo1.ListIndex = 3 Then frmMain.BackColor = vbRed
    If frmMain.Combo1.ListIndex = 4 Then frmMain.BackColor = vbBlue
    If frmMain.Combo1.ListIndex = 5 Then frmMain.BackColor = vbGreen
    If frmMain.Combo1.ListIndex = 6 Then frmMain.BackColor = vbYellow
    If frmMain.Combo1.ListIndex = 7 Then frmMain.BackColor = 16711935
    If frmMain.Combo1.ListIndex = 8 Then frmMain.BackColor = 33023
    If frmMain.Combo1.ListIndex = 9 Then frmMain.BackColor = &H808080
    If frmMain.Combo1.ListIndex = 10 Then frmMain.BackColor = &HFFFF00
End Sub

Public Sub messageuser()
    If messagebox(frmMain.hwnd, "Did this prog actually help?", "help", MB_YESNO Or MB_ICONQUESTION) = IDYES Then GoTo 10 Else GoTo 20
10
    Call messagebox(frmMain.hwnd, "awww....im touched", "i feel special", MB_OK Or MB_ICONASTERISK)
    GoTo 30
20
    Call messagebox(frmMain.hwnd, "ya well fuck you too!", "fuk off", MB_OK Or MB_ICONEXCLAMATION)
30
End Sub

Public Sub setthetimmertime()
    frmMain.Timer3.Enabled = False 'Dissable the timer
    frmMain.Timer3.Interval = "1000" 'Set the interval to 1000
    frmMain.Timer3.Enabled = True 'Enable the timer
End Sub

Public Sub setanothertimer()
    timertext = InputBox("How long do you want the timer to run?", "Timer")
    If timertest > 65535 Then MsgBox "Dude thats too big of a number" 'makes sure that the number is not too big
    frmMain.tiTI.Enabled = False
    frmMain.tiTI.Enabled = True
    frmMain.tiTI.Interval = timertext
    frmMain.Text1.Text = tiTI.Interval
End Sub

Public Sub stopthetimer()
    frmMain.tiTI.Enabled = False
    frmMain.Text1.Text = "Timer stoped" 'Prints out timer stoped in the text box
End Sub

Public Sub opencdtray()
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0) 'Opens the CD tray
End Sub

Public Sub closecdtray()
    retvalue = mciSendString("set CDAudio door close", returnstring, 127, 0) 'Opens the CD tray
End Sub

Public Sub shellinput()
    shelltext = InputBox("Input a shell command", "Shell Command")
    Shell shelltext, vbNormalFocus
End Sub

Public Sub guyzname()
    guy = InputBox("whats your name?", "NAME")
    MsgBox guy & " is gay" 'Opens a message box
End Sub

Public Sub anewdir()
    dirtext = InputBox("Enter directory location and directory name", "Create new Dir") 'Opens a dialog for you to enter text
    CreateNewDirectory (dirtext)
End Sub
