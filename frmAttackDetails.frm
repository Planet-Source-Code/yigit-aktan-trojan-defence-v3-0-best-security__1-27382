VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAttackDetails 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attack details..."
   ClientHeight    =   5040
   ClientLeft      =   2340
   ClientTop       =   825
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Actual Size"
      Height          =   195
      Left            =   5400
      TabIndex        =   26
      Top             =   7200
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Flood Him"
      ForeColor       =   &H8000000F&
      Height          =   1935
      Left            =   720
      MouseIcon       =   "frmAttackDetails.frx":0000
      TabIndex        =   20
      Top             =   5160
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CheckBox chkAddCarriageReturn 
         Caption         =   "Check1"
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   1200
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtFloodTimes 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   960
         Width           =   1365
      End
      Begin VB.TextBox txtFloodMessage 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Bull Shit!"
         Top             =   600
         Width           =   2700
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Enter How Times To Flood:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Enter Your Flood Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2535
      End
      Begin VB.Image picButton 
         Height          =   480
         Left            =   1080
         MousePointer    =   2  'Cross
         Picture         =   "frmAttackDetails.frx":0442
         Top             =   1200
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Attackers IP/ Port/Host"
      ForeColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   750
      TabIndex        =   6
      Top             =   270
      Width           =   3570
      Begin VB.TextBox txtAttckrsIP 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtAttckrsPort 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox txtAttckrsHostName 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   675
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   " IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1575
         TabIndex        =   10
         Top             =   225
         Width           =   390
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Host Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   150
         TabIndex        =   8
         Top             =   675
         Width           =   1140
      End
   End
   Begin VB.TextBox txtAttackingTrojan 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2025
      Width           =   3540
   End
   Begin VB.TextBox txtPortAttacked 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1650
      Width           =   615
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   4920
      Picture         =   "frmAttackDetails.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   990
      Index           =   0
      Left            =   750
      TabIndex        =   3
      Top             =   1425
      Width           =   5340
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Caused by Trojan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   75
         TabIndex        =   5
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port Attacked:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   1590
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   4890
      Left            =   600
      TabIndex        =   13
      Top             =   75
      Width           =   5640
      Begin VB.CommandButton Command4 
         Caption         =   "ShutDown My Computer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         Picture         =   "frmAttackDetails.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Min 
         Caption         =   "Minimize All"
         Height          =   495
         Left            =   4440
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Flood To Attacker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         Picture         =   "frmAttackDetails.frx":37B8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton Logs_save 
         Caption         =   "logs"
         Height          =   195
         Left            =   5280
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "getHost"
         Height          =   195
         Left            =   5280
         TabIndex        =   16
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close My Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Picture         =   "frmAttackDetails.frx":4082
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "OR"
         BeginProperty Font 
            Name            =   "AmericanUncIniD"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Warning !!! "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   7
      Left            =   6300
      Picture         =   "frmAttackDetails.frx":44C4
      Top             =   2100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   6
      Left            =   6300
      Picture         =   "frmAttackDetails.frx":4906
      Top             =   1425
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   5
      Left            =   6300
      Picture         =   "frmAttackDetails.frx":4D48
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   4
      Left            =   6300
      Picture         =   "frmAttackDetails.frx":518A
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   3
      Left            =   0
      Picture         =   "frmAttackDetails.frx":55CC
      Top             =   2100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "frmAttackDetails.frx":5A0E
      Top             =   1425
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "frmAttackDetails.frx":5E50
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "frmAttackDetails.frx":6292
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmAttackDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Const SW_SHOWNORMAL = 1
Const WM_CLOSE = &H10
Const gcClassnameMSWord = "OpusApp"
Const gcClassnameMSExcel = "XLMAIN"
Const gcClassnameMSIExplorer = "IEFrame"
Const gcClassnameMSVBasic = "wndclass_desked_gsk"
Const gcClassnameNotePad = "Notepad"
Const gcClassnameMyVBApp = "ThunderForm"
Dim mintFullHeight As Integer
Dim mintCompactHeight As Integer
Dim mintCurMemoryCounter As Integer
Dim mintTextMemoryCounter As Integer
Dim mstrTextMemory() As String

Private Sub ShowErrorMsg(lngError As Long)
Dim strMessage As String
Select Case lngError
Case WSANOTINITIALISED
strMessage = "A successful WSAStartup call must occur " & _
"before using this function."
Case WSAENETDOWN
strMessage = "The network subsystem has failed."
Case WSAHOST_NOT_FOUND
strMessage = "Authoritative answer host not found."
Case WSATRY_AGAIN
strMessage = "Nonauthoritative host not found, or server failure."
Case WSANO_RECOVERY
strMessage = "A nonrecoverable error occurred."
Case WSANO_DATA
strMessage = "Valid name, no data record of requested type."
Case WSAEINPROGRESS
strMessage = "A blocking Windows Sockets 1.1 call is in " & _
"progress, or the service provider is still " & _
"processing a callback function."
Case WSAEFAULT
strMessage = "The name parameter is not a valid part of " & _
"the user address space."
Case WSAEINTR
strMessage = "A blocking Windows Socket 1.1 call was " & _
"canceled through WSACancelBlockingCall."
End Select
MsgBox strMessage, vbExclamation
End Sub
Public Sub HangUp()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
    lpcConnections)

    If ReturnCode = ERROR_SUCCESS Then
        For i = 0 To lpcConnections - 1
            If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
            End If
        Next i
    End If
End Sub

Public Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    ByteToString = ""
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & Chr(bytString(i))
        i = i + 1
    Wend
End Function




Private Sub cmdGet_Click()
Dim lngInetAdr      As Long
Dim lngPtrHostEnt   As Long
Dim strHostName     As String
Dim udtHostEnt      As HOSTENT
Dim strIpAddress    As String
txtAttckrsHostName.Text = ""
strIpAddress = Trim$(txtAttckrsIP.Text)
lngInetAdr = inet_addr(strIpAddress)
If lngInetAdr = INADDR_NONE Then
ShowErrorMsg (Err.LastDllError)
Else
lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, PF_INET)
If lngPtrHostEnt = 0 Then
ShowErrorMsg (Err.LastDllError)
Else
RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
strHostName = String(256, 0)
RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256
strHostName = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
txtAttckrsHostName.Text = strHostName
End If
End If
End Sub

Private Sub cmdReturn_Click()
MsgBox "Attackers Info saved at C:\Attacker.txt", vbExclamation
Main.cmdStopWatch.Enabled = True
Main.Show
Unload frmAttackDetails
   

   
    
End Sub


Private Sub Command1_Click()
Call HangUp
End Sub

Private Sub Command2_Click()
Me.Height = 7915
End Sub

Private Sub Command3_Click()
Me.Height = 5415
End Sub

Private Sub Command4_Click()
ShutDown_DIALOG
End Sub

Private Sub Logs_save_Click()
' We know Attacker Info, and we must save it (c:\logs.txt)
Dim FSO, Create, NowNow
NowNow = Now
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Create = FSO.CreateTextFile("c:\Attacker.txt")
If txtAttckrsIP.Text = "" = False Then
Create.WriteLine "************************************"
Create.WriteLine "         Trojan Defence v2.0        "
Create.WriteLine "         Warning!  [Attacker]               "
Create.WriteLine "  Attack Day: " & Format(NowNow, "d,mmmm,yyyy ")
Create.WriteLine " Attack Time: " & Format(NowNow, "h:mm:ss ")
Create.WriteLine "   Host Name: " + txtAttckrsHostName.Text + ""
Create.WriteLine "          IP: " + txtAttckrsIP.Text + ""
Create.WriteLine "        Port: " + txtAttckrsPort.Text + ""
Create.WriteLine "************************************"
Create.Close
End If
End Sub

Private Sub Form_Load()



Main.cmdStopWatch.Enabled = False
Min.Value = True
End Sub

Private Function Flood()
'Checks if they type in a room name.
If txtAttckrsIP = "" Then
MsgBox "You can't Flood him. Because, I can't find Attackers IP.", vbExclamation, "Error"
Else
'Sets they variables
Dim strText As String
Dim strTextToSend As String
Dim strChunk As String
Dim intCounter1 As Integer
Dim Time
'Sets time to Zero
Time = 0


'Sets strText to the flood message
        strText = txtFloodMessage.Text

'Checks for special charachters such as { and ]
        For intCounter1 = 1 To Len(strText)
            strChunk = Mid(strText, intCounter1, 1)
            If strChunk = "(" Then strChunk = "{(}"
            If strChunk = ")" Then strChunk = "{)}"
            If strChunk = "+" Then strChunk = "{+}"
            If strChunk = "^" Then strChunk = "{^}"
            If strChunk = "%" Then strChunk = "{%}"
            If strChunk = "~" Then strChunk = "{~}"
            If strChunk = "[" Then strChunk = "{[}"
            If strChunk = "]" Then strChunk = "{]}"
            If strChunk = "{" Then strChunk = "{{}"
            If strChunk = "}" Then strChunk = "{}}"
            strTextToSend = strTextToSend + strChunk
        Next
'Makes it auto hit return
        If chkAddCarriageReturn.Value = Checked Then
            strTextToSend = strTextToSend & Chr(13)
        End If
'Sets the amount of times to flood
For i = 1 To txtFloodTimes.Text
        DoSendKeys txtAttckrsIP.Text, False, strTextToSend, True
'pause amount, which = 0 From Time
s! = Timer
Do: DoEvents
Loop Until Timer - s! > Time
'Checks the array and redims to get rid of unnesasary items
        mintTextMemoryCounter = mintTextMemoryCounter + 1
        ReDim Preserve mstrTextMemory(mintTextMemoryCounter)
        mstrTextMemory(mintTextMemoryCounter) = txtFloodMessage.Text
        mintCurMemoryCounter = UBound(mstrTextMemory) + 1
Next i
      

End If

End Function

Private Sub DoSendKeys(AppToActivate As String, AppActivateDelay As Boolean, TextToSend As String, SendKeysDelay As Boolean)
'This will use SendKeys to send text to an outside application
On Error GoTo ErrHandler

    AppActivate AppToActivate, AppActivateDelay
    SendKeys TextToSend, SendKeysDelay

Exit Sub

ErrHandler:
Exit Sub


End Sub

Private Sub Min_Click()
MinimizeAll
End Sub

Private Sub picButton_Click()
Dim i
'For Flood Progress alitle unreal
ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Max = 300000
For i = 0 To 300000
ProgressBar1.Value = i
Next
ProgressBar1.Visible = False
Call Flood
End Sub
