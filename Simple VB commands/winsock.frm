VERSION 5.00
Begin VB.Form winsock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Winsock :: Simple VB commands :: Arzynik.com"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Text Query"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
      Begin VB.CommandButton Command8 
         Caption         =   "Close Connection"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Go"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "Dissconnect"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtip 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Server IP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label lblremadd 
      Alignment       =   2  'Center
      Caption         =   "Remote Address"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "In order to use this you need to add the winsock componet, mswinsck.ocx  and then add it to the form and name it 'w2'"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   6615
   End
End
Attribute VB_Name = "winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In order to use this you need to add the winsock
'componet, mswinsck.ocx  and then add it to the form
'and name it 'w2'

'________Begin winsock________

Private Sub Command1_Click()
    On Error GoTo err_Command1_Click
    w2.RemotePort = Text2.Text
    If txtip.Text = "" Then
        MsgBox "Enter Computer Name"
    Else
        w2.RemoteHost = txtip.Text
        w2.Connect
    End If
    If w2.State = sckConnected Then
        Me.Caption = "connected to remote host  " & txtip.Text
    Else
        Me.Caption = "not connected to the remote host  " & txtip.Text
    End If
        Exit Sub
err_Command1_Click:
        Screen.MousePointer = vbNormal
        MsgBox "An error has occured." & vbCrLf & vbTab & _
            "Procedure: Command1_Click" & vbCrLf & vbTab & _
            "Error Number: " & Err.Number & vbCrLf & vbTab & _
            "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Command10_Click()
    Dim str As String
    'str = "Chr(&HB2) + Chr(&H0) + Chr(&H0)"
    '67.227.163.70
    str = Text1
    w2.SendData Text1.Text
End Sub

Private Sub Command7_Click()
    Dim str As String
    str = "dir"
    w2.SendData str
    MsgBox "A directory c:\mydir\THIS IS MY POWER\ has been created at " & w2.RemoteHost
End Sub

Private Sub Command2_Click()
    Dim str As String
    str = "closeme"
    w2.SendData str
    w2.Close
    Me.Caption = "Disconnected"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo err_Form_Unload
        w2.Close
    Exit Sub
err_Form_Unload:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Unload" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub w2_Connect()
    On Error GoTo err_w2_Connect
        Me.Caption = "Connected 2 Remote Host"
    Exit Sub
err_w2_Connect:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: w2_Connect" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub w2_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo err_w2_DataArrival
    Dim str As String
    w2.GetData str
    Label2.Caption = str
    Exit Sub
err_w2_DataArrival:
        Screen.MousePointer = vbNormal
        MsgBox "An error has occured." & vbCrLf & vbTab & _
            "Procedure: w2_DataArrival" & vbCrLf & vbTab & _
            "Error Number: " & Err.Number & vbCrLf & vbTab & _
            "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

