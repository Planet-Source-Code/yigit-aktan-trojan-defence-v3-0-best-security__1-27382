VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trojan Defence v3.0"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "Waiting..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdStopWatch 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   765
      Left            =   4125
      Picture         =   "Main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2850
      Width           =   615
   End
   Begin MSWinsockLib.Winsock wskPort 
      Index           =   0
      Left            =   3375
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "N&one"
      Height          =   765
      Left            =   4125
      Picture         =   "Main.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1950
      Width           =   615
   End
   Begin VB.TextBox txtAddPort 
      Height          =   315
      Index           =   1
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4275
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtAddPort 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4275
      Width           =   615
   End
   Begin VB.CommandButton cmdAddPort 
      Caption         =   "Add &port"
      Enabled         =   0   'False
      Height          =   315
      Left            =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&All"
      Height          =   765
      Left            =   4125
      Picture         =   "Main.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   765
      Left            =   4125
      Picture         =   "Main.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3750
      Width           =   615
   End
   Begin VB.CommandButton cmdStartWatch 
      Caption         =   "&Watch"
      Height          =   765
      Left            =   4125
      Picture         =   "Main.frx":111A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   615
   End
   Begin VB.ListBox lstPort 
      BackColor       =   &H80000004&
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
      Height          =   4110
      ItemData        =   "Main.frx":1424
      Left            =   150
      List            =   "Main.frx":142B
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   3690
   End
   Begin VB.Frame Frame1 
      Height          =   4740
      Left            =   75
      TabIndex        =   6
      Top             =   -75
      Width           =   3840
   End
   Begin VB.Frame Frame2 
      Height          =   4740
      Left            =   3975
      TabIndex        =   10
      Top             =   -75
      Width           =   915
   End
   Begin VB.Menu mnuDefence 
      Caption         =   "&Defence"
      Begin VB.Menu mnuWatch 
         Caption         =   "&Watch"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuSystemTray 
      Caption         =   "SysTrayPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Stop watching ports"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Simple Wav Playing Declaration
Private Declare Function sndPlaySound Lib "winmm.dll" Alias _
        "sndPlaySoundA" (ByVal lpszSoundName As String, _
                        ByVal uFlags As Long) As Long

'Boolean used to prevent Form_Activate running more than once
Private blnStartUp As Boolean

Private Sub ExitTrojanDefence()

    'I'm not going to explain this
    Unload Me
    
End Sub


Private Sub GoSystemTray()
    
    'Put application in SysTray and wait for port attack
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hwnd = Me.hwnd
    VBGTray.uId = vbNull
    VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    VBGTray.ucallbackMessage = WM_MOUSEMOVE
    VBGTray.hIcon = Me.Icon
    'tooltiptext
    VBGTray.szTip = Me.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, VBGTray)
    'Hide the main form and enable the SysTray pop-up menu
    Me.Hide
    Me.mnuQuit.Enabled = True
    Me.mnuRestore.Enabled = True
    
End Sub

Private Sub mnuQuit_Click()
    
    'Selected from pop-up menu in SysTray
    StopWatch
    ExitTrojanDefence

End Sub

Private Sub mnuRestore_Click()
    
    'Selected from pop-up menu in SysTray
    'Restore from systray and stop watching ports
    StopWatch
    Me.Show
    'Disable SysTray pop-up menu
    Me.mnuQuit.Enabled = False
    Me.mnuRestore.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lngMsg As Long
    Static blnFlag As Boolean
    Dim result As Long
    
    lngMsg = X / Screen.TwipsPerPixelX


    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
            'doubleclick
            Case WM_LBUTTONDBLCLICK
            'Restore from systray
            StopWatch
            Me.Show
            'Disable SysTray pop-up menu
            Me.mnuQuit.Enabled = False
            Me.mnuRestore.Enabled = False
            'right-click
            Case WM_RBUTTONUP
            result = SetForegroundWindow(Me.hwnd)
            'This menu is on frmTrjDfc2, but hidden
            Me.PopupMenu mnuSystemTray
        End Select
        blnFlag = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    VBGTray.cbSize = Len(VBGTray)
    VBGTray.hwnd = Me.hwnd
    VBGTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
    
End Sub

Private Sub EnableDisable(blnOnOff As Boolean)
    
    'Enable or Disable controls on form
    lstPort.Enabled = blnOnOff
    txtAddPort(0).Enabled = blnOnOff
    cmdAddPort.Enabled = blnOnOff
    cmdClearAll.Enabled = blnOnOff
    cmdExit.Enabled = blnOnOff
    cmdSelectAll.Enabled = blnOnOff
    cmdStartWatch.Enabled = blnOnOff
    mnuDefence.Enabled = blnOnOff
    mnuHelp.Enabled = blnOnOff
    cmdStopWatch.Enabled = Not (blnOnOff)

End Sub

Private Sub AttackDetected(intIndex As Integer)


    'Called from wskPort_ConnectionRequest
    Me.Show
    'Disable SysTray pop-up menu
    Me.mnuQuit.Enabled = False
    Me.mnuRestore.Enabled = False
    'Display report of attack
    frmAttackDetails.Show
    'Populate text boxes on report form using index of winsock control
    frmAttackDetails.txtPortAttacked = Left(lstPort.List(intIndex), 5)
    frmAttackDetails.txtAttackingTrojan = Mid(lstPort.List(intIndex), 8)
    frmAttackDetails.cmdGet.Value = True 'This button for Close your Net connection
    frmAttackDetails.txtAttckrsIP = wskPort(intIndex).RemoteHostIP
    frmAttackDetails.txtAttckrsPort = wskPort(intIndex).RemotePort
    frmAttackDetails.Logs_save.Value = True 'For Attacker info logs Save
    
    'Notify attack with sound
    If Len(App.Path) = 3 Then
        'Running in root directory
        sndPlaySound App.Path & "connect.wav", &H1
    Else
        'Running in subfolder
        sndPlaySound App.Path & "\connect.wav", &H1
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

Private Sub StopWatch()
    Dim intSelect   As Integer
    
    'Loop through all selected ports and close winsock
    For intSelect = 0 To lstPort.ListCount - 1
        If lstPort.Selected(intSelect) Then
            wskPort(intSelect).Close
            While wskPort(intSelect).State
                DoEvents
            Wend
            'If winsock control was created at Runtime then unload it
            If intSelect > 0 Then
                Unload wskPort(intSelect)
            End If
        End If
    Next intSelect
    
    'Re-enable form's controls
    EnableDisable True

End Sub

Private Sub StartWatch()
    Dim intSelect       As Integer
    Dim wskNew          As Winsock
    Dim lngPortCount    As Long
    Dim blnNoWatch      As Boolean
    Dim strNoWatch      As String

    lngPortCount = 0
    strNoWatch = ""
    'Loop through all selected ports and create a new
    'member of the winsock control array to watch that port
    For intSelect = 0 To lstPort.ListCount - 1
        If lstPort.Selected(intSelect) Then
            'reset flag
            blnNoWatch = False
            'Create a new control if more than one port selected
            If intSelect > 0 Then
                Load wskPort(intSelect)
            End If
            wskPort(intSelect).Close
            While wskPort(intSelect).State
                DoEvents
            Wend
            'Set port to watch
            wskPort(intSelect).LocalPort = CLng(Left(lstPort.List(intSelect), 5))
            'The usual cause of an error here is that the
            'port is already in use
            On Error GoTo NoWatch
            wskPort(intSelect).Listen
            'If winsock control initiated ok then keep count
            If Not (blnNoWatch) Then
                lngPortCount = lngPortCount + 1
            End If
        End If
    Next intSelect
            
    'display a list of failed attempts to watch a port
    If strNoWatch <> "" Then
        MsgBox "The following port(s)" & vbCrLf & _
                "cannot be watched:" & vbCrLf & vbCrLf & _
                strNoWatch, _
                vbOKOnly + vbExclamation
    End If
    If lngPortCount > 0 Then
        'Disable form's controls
        EnableDisable False
        StatusBar1.SimpleText = "Watching..."
        MsgBox "If any attack is made" & vbCrLf & _
                "on the selected port(s)," & vbCrLf & _
                "I'll automatically re-open", _
                vbOKOnly + vbInformation, _
                "Watching " & lngPortCount & " port(s)"
        'Put app in SysTray
        GoSystemTray
    Else
        'If start button is clicked when no ports are selected
        'from the list
        MsgBox "No ports are selected!", _
                vbOKOnly + vbExclamation
    
    
    End If

StartWatch_End:
    Exit Sub
    
NoWatch:
    'Appends each failed attempt to listen to a port to a string
    'which is displayed later. blnNoWatch tells the procedure not
    'to count this as a successful port watch
    blnNoWatch = True
    strNoWatch = strNoWatch & wskPort(intSelect).LocalPort & vbCrLf
    Resume Next
    
End Sub

Private Sub cmdStartWatch_Click()
    
    'Obvious
    StartWatch
    StatusBar1.SimpleText = "Watching..."
End Sub

Private Sub cmdStopWatch_Click()

    'Obvious
    StopWatch
    StatusBar1.SimpleText = "Stop Watching..."
    
    
End Sub

Private Sub mnuWatch_Click()
    
    'Obvious
    StartWatch

End Sub


Private Sub cmdExit_Click()

    'Obvious
    ExitTrojanDefence
    
End Sub

Private Sub mnuExit_Click()

    'Obvious
    ExitTrojanDefence
    
End Sub

Private Sub cmdClearAll_Click()
    Dim intSelect   As Integer
    
    'De-select all items in list box
    For intSelect = 0 To lstPort.ListCount - 1
        lstPort.Selected(intSelect) = False
    Next intSelect
    lstPort.Refresh

End Sub

Private Sub cmdSelectAll_Click()
    Dim intSelect   As Integer
    
    'Select all items in list box
    For intSelect = 0 To lstPort.ListCount - 1
        lstPort.Selected(intSelect) = True
    Next intSelect
    lstPort.Refresh
    StatusBar1.SimpleText = "Checked All"
    
    
End Sub

Private Sub ShowErrorMsg(lngError As Long)
    '
    Dim strMessage As String
    '
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
    End Select
    '
    MsgBox strMessage, vbExclamation, "Error..."
    '
End Sub

Private Sub Form_Activate()
    Dim FSO             As FileSystemObject
    Dim txs             As TextStream
    Dim strFile         As String
    Dim strDefinition   As String
    Dim strPortNumber   As String
    Dim lngSplit        As Long
    Dim astrInUse()     As String
    Dim blnInUse        As Boolean
    
    If blnStartUp Then
        ReDim astrInUse(0) As String
        
        'initialise variables
        astrInUse(0) = ""
        cmdAddPort.Enabled = False
        lstPort.Clear
        Set FSO = New FileSystemObject
    
        If Len(App.Path) = 3 Then
            'Running in root directory
            strFile = App.Path & "ports.csv"
        Else
            'Running in Sub-Folder
            strFile = App.Path & "\ports.csv"
        End If
        'Check Reference file exists in program directory
        If FSO.FileExists(strFile) Then
            'Open the file as a textstream
            Set txs = FSO.OpenTextFile(strFile)
            Do Until txs.AtEndOfStream
                'Read file one line at a time
                strDefinition = txs.ReadLine
                'Check for a "," in the line
                lngSplit = InStr(1, strDefinition, ",")
                If lngSplit Then
                    blnInUse = False
                    'Check the port to see if it is already in use
                    wskPort(0).Close
                    While wskPort(0).State
                        DoEvents
                    Wend
                    'Port is the first field in the .csv file, so get it
                    'like this
                    strPortNumber = Left(strDefinition, lngSplit - 1)
                    'Attempt to Listen to port
                    wskPort(0).LocalPort = CLng(strPortNumber)
                    On Error GoTo PortInUse
                    wskPort(0).Listen
                    'If ok then add leading zeros to port number to
                    'make it 5 chars long. Not necessary, but it makes
                    'the list box look tidy. If we need the port
                    'value as Long again then CLng will do the trick
                    If Not (blnInUse) Then
                        Do Until Len(strPortNumber) = 5
                            strPortNumber = "0" & strPortNumber
                        Loop
                        'Add port and trojan name to the list box
                        lstPort.AddItem strPortNumber & vbTab & " " & _
                                        Mid(strDefinition, lngSplit + 1)
                        lstPort.Refresh
                    End If
                End If
            Loop
        Else
            'If the file is not present, a message box is displayed.
            'Much nicer than "Run-Time Error - File Not Found" and allows
            'the program to run, although the user will have to manually
            'add any ports to be watched
            MsgBox "No port definition list found!" & _
                    vbCrLf & vbCrLf & _
                    "Missing: ports.csv" & _
                    vbCrLf & vbCrLf, vbOKOnly + vbInformation, _
                    "File not found..."
        End If
        'An array has been used to store all the port numbers from
        'the reference file that coud not be used.
        If astrInUse(0) <> "" Then
            'This string will now be set ready to display
            'all failed port numbers in the .csv file
            strPortNumber = ""
            'Loop through the array, appending each failed port
            'to the string
            For lngSplit = LBound(astrInUse) To UBound(astrInUse)
                strPortNumber = strPortNumber & vbCrLf & astrInUse(lngSplit)
            Next lngSplit
            'Redim array to save memory (not a lot)
            ReDim astrInUse(0) As String
            'See if more than one port number failed.
            'A more general msgbox could be used, as elsewhere
            'in the code, but it just looks nice this way.
            'I'm just a bit old fashioned i suppose :o)
            If lngSplit > 2 Then
                MsgBox "The following ports" & vbCrLf & _
                        "are in use and have" & vbCrLf & _
                        "not been loaded" & vbCrLf & _
                        "from file:" & vbCrLf & vbCrLf & _
                        strPortNumber & vbCrLf & vbCrLf, _
                        vbOKOnly + vbInformation, _
                        "Ports in use..."
            Else
                MsgBox "The following port" & vbCrLf & _
                        "is in use and has" & vbCrLf & _
                        "not been loaded" & vbCrLf & _
                        "from file:" & vbCrLf & vbCrLf & _
                        strPortNumber & vbCrLf & vbCrLf, _
                        vbOKOnly + vbInformation, _
                        "Port in use..."
            End If
        End If
        
        'Clean up
        Set txs = Nothing
        Set FSO = Nothing
        blnStartUp = False
    End If
    
Form_Activate_End:
    Exit Sub
    
PortInUse:
    'If there is an error when trying to listen to a port
    '(normally because it is already in use), we add the
    'relevant port number to the array
    blnInUse = True
    astrInUse(UBound(astrInUse)) = strPortNumber
    ReDim Preserve astrInUse(UBound(astrInUse) + 1) As String
    Resume Next

End Sub

Private Sub Form_Load()
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
 
    'Set the one-time flag
    blnStartUp = True
    'Check for valid Winsock enviroment
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    If lngRetVal <> 0 Then
        Select Case lngRetVal
        Case WSASYSNOTREADY
            strErrorMsg = "The underlying network subsystem is not " & _
                "ready for network communication."
        Case WSAVERNOTSUPPORTED
            strErrorMsg = "The version of Windows Sockets API support " & _
                "requested is not provided by this particular " & _
                "Windows Sockets implementation."
        Case WSAEINVAL
            strErrorMsg = "The Windows Sockets version specified by the " & _
                "application is not supported by this DLL."
        End Select
        'Report any error
        MsgBox strErrorMsg, vbCritical
            End If
    
End Sub

Private Sub mnuAbout_Click()
    
   aBOUT.Show
   
   
   
   
   
End Sub

Private Sub txtAddPort_Change(Index As Integer)
    
    'Enables the 'Add Port to List' button if there is
    'anything in the 'Port Number to Add' text box
    'Sets the description to "User selected"
    If txtAddPort(0) = "" Then
        cmdAddPort.Enabled = False
    Else
        cmdAddPort.Enabled = True
        txtAddPort(1) = "User selected port"
    End If
    
End Sub

Private Sub txtAddPort_LostFocus(Index As Integer)

    'Horrible bit of code to add leading zeros to the
    'port number before it is added to the list. Also
    'makes sure it is a number.
    If txtAddPort(0) Like "#####" Then
        txtAddPort(0) = txtAddPort(0)
        cmdAddPort.SetFocus
    ElseIf txtAddPort(0) Like "####" Then
        txtAddPort(0) = "0" & txtAddPort(0)
        cmdAddPort.SetFocus
    ElseIf txtAddPort(0) Like "###" Then
        txtAddPort(0) = "00" & txtAddPort(0)
        cmdAddPort.SetFocus
    ElseIf txtAddPort(0) Like "##" Then
        txtAddPort(0) = "000" & txtAddPort(0)
        cmdAddPort.SetFocus
    ElseIf txtAddPort(0) Like "#" Then
        txtAddPort(0) = "0000" & txtAddPort(0)
        cmdAddPort.SetFocus
    ElseIf txtAddPort(0) Like "?*" Then
        MsgBox "Not a valid port!", vbOKOnly + vbExclamation
        txtAddPort(0) = ""
        txtAddPort(1) = ""
        txtAddPort(0).SetFocus
    End If
    
End Sub

Private Sub cmdAddPort_Click()
    Dim blnInUse    As Boolean
    
    'Before adding the port number to the list, a quick
    'check is made to see if the port is already in use.
    blnInUse = False
    wskPort(0).Close
    While wskPort(0).State
        DoEvents
    Wend
    wskPort(0).LocalPort = txtAddPort(0)
    On Error GoTo CannotAddToList
    wskPort(0).Listen
    If Not (blnInUse) Then
        'If all is ok then add the new port number to
        'the list box
        lstPort.AddItem txtAddPort(0) & vbTab & " " & txtAddPort(1), lstPort.ListCount
        StatusBar1.SimpleText = "Added..."
        'and select it. An assumption here that if a user has
        'manually added a port number, then they might want it
        'to be watched. A more permanent solution is to add the
        'port and description to ports.csv which should be in
        'the same directory as the program.
        lstPort.Selected(lstPort.ListCount - 1) = True
        lstPort.Refresh
    End If
    'reset the text boxes
    txtAddPort(0) = ""
    txtAddPort(1) = ""

cmdAddPort_Click_End:
    Exit Sub
    
CannotAddToList:
    blnInUse = True
    'if there is an error, this is reported
    MsgBox "Port " & txtAddPort(0) & " is already in use" & vbCrLf & _
            "and will not be added to the" & vbCrLf & _
            "list of ports to watch.", _
            vbOKOnly + vbInformation, _
            "Port in use..."
    Resume Next
    
End Sub

Private Sub wskPort_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    'If a connection attempt is made on any of the
    'watched ports then the appropriate winsock control
    'fires this event. Because it a member of an array
    'we can use it's index value to find out which port
    'was attacked
    AttackDetected Index

End Sub

'That's all folks
'Have fun
'Yigit Aktan / M@verick
