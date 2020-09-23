VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox CommandBar 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2280
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox DisplayBox 
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mana"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Energy"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1560
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2040
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1440
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnectToggle 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFileMacros 
         Caption         =   "&Macros"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuFileAliases 
         Caption         =   "&Aliases"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandBar_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LoadError
    'below are the macro keys
    If KeyCode = vbKeyF5 Then
        KeySection = "Macros"
        KeyKey = "F5"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF6 Then
        KeySection = "Macros"
        KeyKey = "F6"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF7 Then
        KeySection = "Macros"
        KeyKey = "F7"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF8 Then
        KeySection = "Macros"
        KeyKey = "F8"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF9 Then
        KeySection = "Macros"
        KeyKey = "F9"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF10 Then
        KeySection = "Macros"
        KeyKey = "F10"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF11 Then
        KeySection = "Macros"
        KeyKey = "F11"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
    If KeyCode = vbKeyF12 Then
        KeySection = "Macros"
        KeyKey = "F12"
        LoadINI
        CommandBar.Text = CommandBar.Text & KeyValue
        CommandBar.SelStart = Len(CommandBar.Text)
        CommandBar.SelLength = 0
        KeyCode = 0
    End If
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            MsgBox "This macro is empty."
    End Select
End Sub

Private Sub CommandBar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyNumpad8 Then ' Or KeyCode = 38 go north
        CommandBar.Text = "n"
        CommandBar_KeyPress 13
    End If
    If KeyCode = vbKeyNumpad2 Then ' Or KeyCode = 40 go south
        CommandBar.Text = "s"
        CommandBar_KeyPress 13
    End If
    If KeyCode = vbKeyNumpad4 Then ' Or KeyCode = 37 go west
        CommandBar.Text = "w"
        CommandBar_KeyPress 13
    End If
    If KeyCode = vbKeyNumpad6 Then ' Or KeyCode = 39 go east
        CommandBar.Text = "e"
        CommandBar_KeyPress 13
    End If
End Sub

Private Sub Form_Load()
    LastColor = "3"
    'highlight display text
    DisplayBox.SelStart = 0
    DisplayBox.SelLength = Len(DisplayBox.Text)
    'turn text white
    DisplayBox.SelColor = vbWhite
    'remove highlight from text
    DisplayBox.SelLength = 0
    'Load Option settings
    'take care of missing ini file error
    On Error GoTo LoadError
    KeySection = "Options"
    KeyKey = "Alias"
    LoadINI
    If InStr(1, KeyValue, "True") Then
        frmOptions.chkAlias.Caption = "Alias On"
        frmOptions.chkAlias.Value = 1
    ElseIf InStr(1, KeyValue, "False") Then
        frmOptions.chkAlias.Caption = "Alias Off"
        frmOptions.chkAlias.Value = 0
    End If
    KeySection = "Options"
    KeyKey = "Echo"
    LoadINI
    If InStr(1, KeyValue, "True") Then
        frmOptions.chkEcho.Caption = "Echo On (repeats to you what you enter)"
        frmOptions.chkEcho.Value = 1
    ElseIf InStr(1, KeyValue, "False") Then
        frmOptions.chkEcho.Caption = "Echo Off (repeats to you what you enter)"
        frmOptions.chkEcho.Value = 0
    End If
    KeySection = "Options"
    KeyKey = "Size"
    LoadINI
    Form_Resize
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            'do nothing
    End Select
End Sub

Private Sub Form_Resize() 'make sure everything looks right no matter what shape the form is in
    On Error Resume Next
    DisplayBox.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 900
    DisplayBox.RightMargin = DisplayBox.Width - 400
    CommandBar.Move 100, Me.Height - 1400, Me.ScaleWidth - 200
    Dim Temp As Long
    Temp = (Me.Width - 300) / 3
    Shape1.Move 100, Me.Height - 995, Me.ScaleWidth - ((Temp * 2) + 200)
    Shape2.Move 100 + (Me.ScaleWidth - ((Temp * 2) + 200)), Me.Height - 995, Me.ScaleWidth - ((Temp * 2) + 200)
    Shape3.Move 100 + (Me.ScaleWidth - ((Temp * 2) + 200)) + (Me.ScaleWidth - ((Temp * 2) + 200)), Me.Height - 995, Me.ScaleWidth - ((Temp * 2) + 200)
    Label1.Move 100, Me.Height - 995 + 50, Me.ScaleWidth - ((Temp * 2) + 200)
    Label2.Move 100 + (Me.ScaleWidth - ((Temp * 2) + 200)), Me.Height - 995 + 50, Me.ScaleWidth - ((Temp * 2) + 200)
    Label3.Move 100 + (Me.ScaleWidth - ((Temp * 2) + 200)) + (Me.ScaleWidth - ((Temp * 2) + 200)), Me.Height - 995 + 50, Me.ScaleWidth - ((Temp * 2) + 200)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmOptions.chkSize.Value = 0 Then
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "False"
        SaveINI
    ElseIf frmOptions.chkSize.Value = 1 Then
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "True"
        SaveINI
        KeySection = "Size"
        KeyKey = "Height"
        KeyValue = Me.Height
        SaveINI
        KeySection = "Size"
        KeyKey = "Width"
        KeyValue = Me.Width
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleHeight"
        KeyValue = Me.ScaleHeight
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleWidth"
        KeyValue = Me.ScaleWidth
        SaveINI
    End If
    End
End Sub

Private Sub CommandBar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'if you hit Enter then
        If Winsock1.State = 7 Then 'check to see if connected
            If frmOptions.chkAlias.Value = 1 Then 'if aliases are turned on
                AliasCheck CommandBar.Text
                If frmOptions.chkEcho.Value = 1 Then Display YELLOW & CommandBar.Text & vbCrLf & bYELLOW
                CommandBar.SelStart = 0
                CommandBar.SelLength = Len(CommandBar.Text)
                If endIt Then Exit Sub
            Else
                If frmOptions.chkEcho.Value = 1 Then Display YELLOW & CommandBar.Text & vbCrLf & bYELLOW
            End If
            Winsock1.SendData CommandBar.Text & vbCrLf
        End If
        'highlight the text
        CommandBar.SelStart = 0
        CommandBar.SelLength = Len(CommandBar.Text)
        If KeyAscii = 13 Then KeyAscii = 0 'this way it doesn't beep when you hit enter
    End If
End Sub

Private Sub mnuFileAliases_Click()
    frmAliases.Show
End Sub

Private Sub mnuFileConnectToggle_Click()
    If mnuFileConnectToggle.Caption = "&Connect" Then
        frmConnect.Show
        Me.Hide
    ElseIf mnuFileConnectToggle.Caption = "&Disconnect" Then
        Winsock1.Close
        mnuFileConnectToggle.Caption = "&Connect"
        Me.Caption = frmConnect.txtName.Text & " ::Disconnected"
        Dim tmpDisplay As String
        tmpDisplay = WHITE & vbCrLf & "***Disconnected***" & vbCrLf & GREEN
        Display tmpDisplay
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMacros_Click()
    frmMacros.Show
End Sub

Private Sub mnuFileOptions_Click()
    frmOptions.Show
End Sub

Private Sub Winsock1_Close()
    mnuFileConnectToggle.Caption = "&Connect"
    Me.Caption = frmConnect.txtName.Text & " ::Disconnected"
    Dim tmpDisplay As String
    tmpDisplay = WHITE & vbCrLf & "***Disconnected***" & vbCrLf & GREEN
    Display tmpDisplay
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    mnuFileConnectToggle.Caption = "&Disconnect"
    Me.Caption = frmConnect.txtName.Text & " ::Connected"
    Dim tmpDisplay As String
    tmpDisplay = WHITE & vbCrLf & "Connected to " & Winsock1.RemoteHost & GREEN
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim UserCommand As String
    Dim WrkSpace As String
    Winsock1.GetData UserCommand
    'Clipboard.SetText UserCommand
    If InStr(1, UserCommand, "[2J") Then 'this whole if thing takes care of the crap sometimes attached to the beginning of a message
        WrkSpace = Mid(UserCommand, 5, Len(UserCommand) - 5)
        Display WrkSpace
        Exit Sub
    End If
    'if no crap is in front of it display it as is
    Display UserCommand
End Sub

Private Sub Display(WrkSpace As String)
    Dim strTemp As String
    Dim WrkSpace2 As String
    Dim InTemp As Integer
    Dim LastBLACK As Integer
    Dim LastRED As Integer
    Dim LastGREEN As Integer
    Dim LastYELLOW As Integer
    Dim LastBLUE As Integer
    Dim LastMAGNETA As Integer
    Dim LastLIGHTBLUE As Integer
    Dim LastWHITE As Integer
    Dim LastbBLACK As Integer
    Dim LastbRED As Integer
    Dim LastbGREEN As Integer
    Dim LastbYELLOW As Integer
    Dim LastbBLUE As Integer
    Dim LastbMAGNETA As Integer
    Dim LastbLIGHTBLUE As Integer
    Dim LastbWHITE As Integer
    MsgBox WrkSpace
    'find the last of each color in the message
    WrkSpace2 = WrkSpace
    strTemp = "[1m[30m"
    InTemp = 1
    LastBLACK = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastBLACK = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
    Loop
    strTemp = "[1m[31m"
    InTemp = 1
    LastRED = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastRED = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[32m"
    InTemp = 1
    LastGREEN = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastGREEN = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[33m"
    InTemp = 1
    LastYELLOW = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastYELLOW = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[34m"
    InTemp = 1
    LastBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[35m"
    InTemp = 1
    LastMAGNETA = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastMAGNETA = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[36m"
    InTemp = 1
    LastLIGHTBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastLIGHTBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[1m[37m"
    InTemp = 1
    LastWHITE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastWHITE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[30m"
    InTemp = 1
    LastbBLACK = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbBLACK = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[31m"
    InTemp = 1
    LastbRED = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbRED = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[32m"
    InTemp = 1
    LastbGREEN = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbGREEN = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[33m"
    InTemp = 1
    LastbYELLOW = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbYELLOW = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[34m"
    InTemp = 1
    LastbBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[35m"
    InTemp = 1
    LastbMAGNETA = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbMAGNETA = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[36m"
    InTemp = 1
    LastbLIGHTBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbLIGHTBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    strTemp = "[0m[37m"
    InTemp = 1
    LastbWHITE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0
        LastbWHITE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then
            InTemp = Len(WrkSpace2)
        End If
    Loop
    'compare to see which color is the last one
    If LastBLACK > LastRED & LastBLACK > LastGREEN & LastBLACK > LastYELLOW & LastBLACK > LastBLUE & LastBLACK > LastMAGNETA & LastBLACK > LastLIGHTBLUE & LastBLACK > LastWHITE & LastBLACK > LastbBLACK & LastBLACK > LastbRED & LastBLACK > LastbGREEN & LastBLACK > LastbYELLOW & LastBLACK > LastbBLUE & LastBLACK > LastbMAGNETA & LastBLACK > LastbLIGHTBLUE & LastBLACK > LastbWHITE Then
        LastColor = "0"
    ElseIf LastRED > LastBLACK & LastRED > LastGREEN & LastRED > LastYELLOW & LastRED > LastBLUE & LastRED > LastMAGNETA & LastRED > LastLIGHTBLUE & LastRED > LastWHITE & LastRED > LastbBLACK & LastRED > LastbRED & LastRED > LastbGREEN & LastRED > LastbYELLOW & LastRED > LastbBLUE & LastRED > LastbMAGNETA & LastRED > LastbLIGHTBLUE & LastRED > LastbWHITE Then
        LastColor = "2"
    ElseIf LastGREEN > LastBLACK & LastGREEN > LastRED & LastGREEN > LastYELLOW & LastGREEN > LastBLUE & LastGREEN > LastMAGNETA & LastGREEN > LastLIGHTBLUE & LastGREEN > LastWHITE & LastGREEN > LastbBLACK & LastGREEN > LastbRED & LastGREEN > LastbGREEN & LastGREEN > LastbYELLOW & LastGREEN > LastbBLUE & LastGREEN > LastbMAGNETA & LastGREEN > LastbLIGHTBLUE & LastGREEN > LastbWHITE Then
        LastColor = "3"
    ElseIf LastYELLOW > LastBLACK & LastYELLOW > LastRED & LastYELLOW > LastGREEN & LastYELLOW > LastBLUE & LastYELLOW > LastMAGNETA & LastYELLOW > LastLIGHTBLUE & LastYELLOW > LastWHITE & LastYELLOW > LastbBLACK & LastYELLOW > LastbRED & LastYELLOW > LastbGREEN & LastYELLOW > LastbYELLOW & LastYELLOW > LastbBLUE & LastYELLOW > LastbMAGNETA & LastYELLOW > LastbLIGHTBLUE & LastYELLOW > LastbWHITE Then
        LastColor = "4"
    ElseIf LastBLUE > LastBLACK & LastBLUE > LastRED & LastBLUE > LastGREEN & LastBLUE > LastYELLOW & LastBLUE > LastMAGNETA & LastBLUE > LastLIGHTBLUE & LastBLUE > LastWHITE & LastBLUE > LastbBLACK & LastBLUE > LastbRED & LastBLUE > LastbGREEN & LastBLUE > LastbYELLOW & LastBLUE > LastbBLUE & LastBLUE > LastbMAGNETA & LastBLUE > LastbLIGHTBLUE & LastBLUE > LastbWHITE Then
        LastColor = "1"
    ElseIf LastMAGNETA > LastBLACK & LastMAGNETA > LastRED & LastMAGNETA > LastGREEN & LastMAGNETA > LastYELLOW & LastMAGNETA > LastBLUE & LastMAGNETA > LastLIGHTBLUE & LastMAGNETA > LastWHITE & LastMAGNETA > LastbBLACK & LastMAGNETA > LastbRED & LastMAGNETA > LastbGREEN & LastMAGNETA > LastbYELLOW & LastMAGNETA > LastbBLUE & LastMAGNETA > LastbMAGNETA & LastMAGNETA > LastbLIGHTBLUE & LastMAGNETA > LastbWHITE Then
        LastColor = "7"
    ElseIf LastLIGHTBLUE > LastBLACK & LastLIGHTBLUE > LastRED & LastLIGHTBLUE > LastGREEN & LastLIGHTBLUE > LastYELLOW & LastLIGHTBLUE > LastBLUE & LastLIGHTBLUE > LastMAGNETA & LastLIGHTBLUE > LastWHITE & LastLIGHTBLUE > LastbBLACK & LastLIGHTBLUE > LastbRED & LastLIGHTBLUE > LastbGREEN & LastLIGHTBLUE > LastbYELLOW & LastLIGHTBLUE > LastbBLUE & LastLIGHTBLUE > LastbMAGNETA & LastLIGHTBLUE > LastbLIGHTBLUE & LastLIGHTBLUE > LastbWHITE Then
        LastColor = "5"
    ElseIf LastWHITE > LastBLACK & LastWHITE > LastRED & LastWHITE > LastGREEN & LastWHITE > LastYELLOW & LastWHITE > LastBLUE & LastWHITE > LastMAGNETA & LastWHITE > LastLIGHTBLUE & LastWHITE > LastbBLACK & LastWHITE > LastbRED & LastWHITE > LastbGREEN & LastWHITE > LastbYELLOW & LastWHITE > LastbBLUE & LastWHITE > LastbMAGNETA & LastWHITE > LastbLIGHTBLUE & LastWHITE > LastbWHITE Then
        LastColor = "6"
    ElseIf LastbBLACK > LastRED & LastbBLACK > LastGREEN & LastbBLACK > LastYELLOW & LastbBLACK > LastBLUE & LastbBLACK > LastMAGNETA & LastbBLACK > LastLIGHTBLUE & LastbBLACK > LastWHITE & LastbBLACK > LastBLACK & LastbBLACK > LastbRED & LastbBLACK > LastbGREEN & LastbBLACK > LastbYELLOW & LastbBLACK > LastbBLUE & LastbBLACK > LastbMAGNETA & LastbBLACK > LastbLIGHTBLUE & LastbBLACK > LastbWHITE Then
        LastColor = "0"
    ElseIf LastbRED > LastBLACK & LastbRED > LastGREEN & LastbRED > LastYELLOW & LastbRED > LastBLUE & LastbRED > LastMAGNETA & LastbRED > LastLIGHTBLUE & LastbRED > LastWHITE & LastbRED > LastbBLACK & LastbRED > LastRED & LastbRED > LastbGREEN & LastbRED > LastbYELLOW & LastbRED > LastbBLUE & LastbRED > LastbMAGNETA & LastbRED > LastbLIGHTBLUE & LastbRED > LastbWHITE Then
        LastColor = "10"
    ElseIf LastbGREEN > LastBLACK & LastbGREEN > LastRED & LastbGREEN > LastYELLOW & LastbGREEN > LastBLUE & LastbGREEN > LastMAGNETA & LastbGREEN > LastLIGHTBLUE & LastbGREEN > LastWHITE & LastbGREEN > LastbBLACK & LastbGREEN > LastbRED & LastbGREEN > LastGREEN & LastbGREEN > LastbYELLOW & LastbGREEN > LastbBLUE & LastbGREEN > LastbMAGNETA & LastbGREEN > LastbLIGHTBLUE & LastbGREEN > LastbWHITE Then
        LastColor = "9"
    ElseIf LastbYELLOW > LastBLACK & LastbYELLOW > LastRED & LastbYELLOW > LastGREEN & LastbYELLOW > LastBLUE & LastbYELLOW > LastMAGNETA & LastbYELLOW > LastLIGHTBLUE & LastbYELLOW > LastWHITE & LastbYELLOW > LastbBLACK & LastbYELLOW > LastbRED & LastbYELLOW > LastbGREEN & LastbYELLOW > LastYELLOW & LastbYELLOW > LastbBLUE & LastbYELLOW > LastbMAGNETA & LastbYELLOW > LastbLIGHTBLUE & LastbYELLOW > LastbWHITE Then
        LastColor = "11"
    ElseIf LastbBLUE > LastBLACK & LastbBLUE > LastRED & LastbBLUE > LastGREEN & LastbBLUE > LastYELLOW & LastbBLUE > LastMAGNETA & LastbBLUE > LastLIGHTBLUE & LastbBLUE > LastWHITE & LastbBLUE > LastbBLACK & LastbBLUE > LastbRED & LastbBLUE > LastbGREEN & LastbBLUE > LastbYELLOW & LastbBLUE > LastBLUE & LastbBLUE > LastbMAGNETA & LastbBLUE > LastbLIGHTBLUE & LastbBLUE > LastbWHITE Then
        LastColor = "12"
    ElseIf LastbMAGNETA > LastBLACK & LastbMAGNETA > LastRED & LastbMAGNETA > LastGREEN & LastbMAGNETA > LastYELLOW & LastbMAGNETA > LastBLUE & LastbMAGNETA > LastLIGHTBLUE & LastbMAGNETA > LastWHITE & LastbMAGNETA > LastbBLACK & LastbMAGNETA > LastbRED & LastbMAGNETA > LastbGREEN & LastbMAGNETA > LastbYELLOW & LastbMAGNETA > LastbBLUE & LastbMAGNETA > LastbMAGNETA & LastbMAGNETA > LastbLIGHTBLUE & LastbMAGNETA > LastbWHITE Then
        LastColor = "14"
    ElseIf LastbLIGHTBLUE > LastBLACK & LastbLIGHTBLUE > LastRED & LastbLIGHTBLUE > LastGREEN & LastbLIGHTBLUE > LastYELLOW & LastbLIGHTBLUE > LastBLUE & LastbLIGHTBLUE > LastMAGNETA & LastbLIGHTBLUE > LastWHITE & LastbLIGHTBLUE > LastbBLACK & LastbLIGHTBLUE > LastbRED & LastbLIGHTBLUE > LastbGREEN & LastbLIGHTBLUE > LastbYELLOW & LastbLIGHTBLUE > LastbBLUE & LastbLIGHTBLUE > LastbMAGNETA & LastbLIGHTBLUE > LastLIGHTBLUE & LastbLIGHTBLUE > LastbWHITE Then
        LastColor = "15"
    ElseIf LastbWHITE > LastBLACK & LastbWHITE > LastRED & LastbWHITE > LastGREEN & LastbWHITE > LastYELLOW & LastbWHITE > LastBLUE & LastbWHITE > LastMAGNETA & LastbWHITE > LastLIGHTBLUE & LastbWHITE > LastbBLACK & LastbWHITE > LastbRED & LastbWHITE > LastbGREEN & LastbWHITE > LastbYELLOW & LastbWHITE > LastbBLUE & LastbWHITE > LastbMAGNETA & LastbWHITE > LastbLIGHTBLUE & LastbWHITE > LastWHITE Then
        LastColor = "8"
    End If
    Dim TempBegin As Integer
    Dim TempStrLength As Integer
    Dim TempDebug As String
    TempBegin = 1
    If InStr(1, WrkSpace, "37m") Then
        Do While InStr(TempBegin, WrkSpace, "[1") > 0
            TempBegin = InStr(TempBegin + 1, WrkSpace, "[1")
            If TempBegin > InStr(1, WrkSpace, "37m") Then Exit Do
            TempStrLength = Len(Mid(WrkSpace, TempBegin, InStr(1, WrkSpace, "37m") - TempBegin))
            'MsgBox TempDebug = Mid(WrkSpace, TempBegin, InStr(1, WrkSpace, "30m") - TempBegin)
            If TempStrLength < 11 & TempStrLength > 5 Then
                WHITE = Mid(WrkSpace, TempBegin, InStr(1, WrkSpace, "37m") - TempBegin)
                GoTo ContinueWHITE
            End If
        Loop
        TempBegin = 1
        Do While InStr(TempBegin, WrkSpace, "37m") > 0
            TempBegin = InStr(TempBegin + 1, WrkSpace, "[1")
            If TempBegin < InStr(1, WrkSpace, "37m") Then Exit Do
            TempStrLength = Len(Mid(WrkSpace, TempBegin, InStr(1, WrkSpace, "[1") - TempBegin))
            If TempStrLength > -11 & TempStrLength < -5 Then
                WHITE = Mid(WrkSpace, TempBegin, InStr(1, wrkspack, "[1") - TempBegin)
                GoTo ContinueWHITE
            End If
        Loop
        'GoTo beginred
        GoTo ContinueWHITE
    End If
ContinueWHITE:
    'Format the message so it can be fed into the display box's RTF
    WrkSpace = Replace(WrkSpace, "\", "\\")     '
    WrkSpace = Replace(WrkSpace, "{", "\{")     'these three replace stuff that could be a monkey wrench to the code
    WrkSpace = Replace(WrkSpace, "}", "\}")     '
    WrkSpace = Replace(WrkSpace, vbCr, vbCrLf)              'takes care of returns ... just not sure as to how well though
    WrkSpace = Replace(WrkSpace, vbLg, vbCrLf)              '
    WrkSpace = Replace(WrkSpace, vbCrLf, vbCrLf & "\par ")  '
    WrkSpace = Replace(WrkSpace, "ÿû", " ")    'gets rid of some garbage off normal servers
    WrkSpace = Replace(WrkSpace, "ÿü", vbCrLf) '
    WrkSpace = Replace(WrkSpace, BLACK, "\plain\f2\fs20\cf0 ")          'Begin color
    WrkSpace = Replace(WrkSpace, BLUE, "\plain\f2\fs20\cf1 ")           'these replace the crap color codes with useful ones
    WrkSpace = Replace(WrkSpace, RED, "\plain\f2\fs20\cf2 ")            '
    WrkSpace = Replace(WrkSpace, GREEN, "\plain\f2\fs20\cf3 ")          '
    WrkSpace = Replace(WrkSpace, YELLOW, "\plain\f2\fs20\cf4 ")         '
    WrkSpace = Replace(WrkSpace, MAGNETA, "\plain\f2\fs20\cf7 ")        '
    WrkSpace = Replace(WrkSpace, LIGHTBLUE, "\plain\f2\fs20\cf5 ") '
    WrkSpace = Replace(WrkSpace, WHITE, "\plain\f2\fs20\cf6 ")          '
    WrkSpace = Replace(WrkSpace, bBLACK, "\plain\f2\fs20\cf0 ")         'End color (are all white)
    WrkSpace = Replace(WrkSpace, bBLUE, "\plain\f2\fs20\cf12 ")         '
    WrkSpace = Replace(WrkSpace, bRED, "\plain\f2\fs20\cf10 ")          '
    WrkSpace = Replace(WrkSpace, bGREEN, "\plain\f2\fs20\cf9 ")         '
    WrkSpace = Replace(WrkSpace, bYELLOW, "\plain\f2\fs20\cf11 ")       '
    WrkSpace = Replace(WrkSpace, bMAGNETA, "\plain\f2\fs20\cf14 ")      '
    WrkSpace = Replace(WrkSpace, bLIGHTBLUE, "\plain\f2\fs20\cf15 ")    '
    WrkSpace = Replace(WrkSpace, bWHITE, "\plain\f2\fs20\cf8 ")         '
    'this filters colors for other servers
    WrkSpace = Replace(WrkSpace, BLACK2, "\plain\f2\fs20\cf0 ")         'Begin color2
    WrkSpace = Replace(WrkSpace, BLUE2, "\plain\f2\fs20\cf1 ")          'these replace the crap color codes with useful ones
    WrkSpace = Replace(WrkSpace, RED2, "\plain\f2\fs20\cf2 ")           '
    WrkSpace = Replace(WrkSpace, GREEN2, "\plain\f2\fs20\cf3 ")         '
    WrkSpace = Replace(WrkSpace, YELLOW2, "\plain\f2\fs20\cf4 ")        '
    WrkSpace = Replace(WrkSpace, MAGNETA2, "\plain\f2\fs20\cf7 ")       '
    WrkSpace = Replace(WrkSpace, LIGHTBLUE2, "\plain\f2\fs20\cf5 ")     '
    'WrkSpace = Replace(WrkSpace, WHITE2, "\plain\f2\fs20\cf6 ")         '
    WrkSpace = Replace(WrkSpace, bBLACK2, "\plain\f2\fs20\cf0 ")        'End color2 (are all white)
    WrkSpace = Replace(WrkSpace, bBLUE2, "\plain\f2\fs20\cf12 ")        '
    WrkSpace = Replace(WrkSpace, bRED2, "\plain\f2\fs20\cf10 ")         '
    WrkSpace = Replace(WrkSpace, bGREEN2, "\plain\f2\fs20\cf9 ")        '
    WrkSpace = Replace(WrkSpace, bYELLOW2, "\plain\f2\fs20\cf11 ")      '
    WrkSpace = Replace(WrkSpace, bMAGNETA2, "\plain\f2\fs20\cf14 ")     '
    WrkSpace = Replace(WrkSpace, bLIGHTBLUE2, "\plain\f2\fs20\cf15 ")   '
    WrkSpace = Replace(WrkSpace, bWHITE2, "\plain\f2\fs20\cf8 ")        '
    WrkSpace = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fmodern OCR A Extended;}}" & vbCrLf & "{\colortbl\red127\green127\blue127;\red0\green0\blue255;\red255\green0\blue0;\red0\green255\blue0;\red255\green255\blue0;\red128\green255\blue255;\red255\green255\blue255;\red255\green0\blue175;\red200\green200\blue200;\red0\green125\blue0;\red150\green0\blue0;\red125\green125\blue0;\red0\green0\blue125;\red0\green0\blue0;\red151\green0\blue102;\red0\green150\blue150;}" & vbCrLf & "\deflang1033\pard\plain\f2\fs20\cf" & LastColor & " " & WrkSpace  'adding stuff to the beginning so it will work right
    WrkSpace = WrkSpace & vbCrLf & "\plain\f2\fs20\par }"   'adding stuff to the end so it'll work right
    Clipboard.SetText WrkSpace
    'put cursor at the end of the text
    DisplayBox.SelStart = Len(DisplayBox.Text)
    DisplayBox.SelLength = 0
    'input the message
    DisplayBox.SelRTF = WrkSpace
    'scroll down to the end to show what's new
    DisplayBox.SelStart = Len(DisplayBox.Text)
    DisplayBox.SelLength = 0
End Sub

Private Sub LoadINI()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\Settings.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
'Call MsgBox("An error has occurred while calling the API function", vbExclamation)
KeyValue = ""
Else
KeyValue = Trim(strResult)
End If

End Sub

Private Sub SaveINI()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\Settings.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub

Private Sub LoadAINI()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\AliasList.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub

Private Sub AliasCheck(Message As String)
    endIt = False
    KeySection = "Count"
    KeyKey = "AliasCount"
    LoadAINI
    Dim AliasCount As String
    AliasCount = KeyValue
    Dim strTemp As String
    For a = 1 To Val(AliasCount)
        KeySection = "Aliases"
        KeyKey = "Alias" & a
        LoadAINI
        frmAliases.txtAlias.Text = KeyValue
        KeySection = "Commands"
        KeyKey = "Command" & a
        LoadAINI
        frmAliases.txtCommand.Text = KeyValue
        If frmAliases.txtAlias.Text = MidWord(Message, 1, Words(frmAliases.txtAlias.Text)) Then
            strTemp = Replace(Message, Mid(Message, 1, Len(frmAliases.txtAlias.Text)), frmAliases.txtCommand.Text, 1, 1)
            Winsock1.SendData strTemp & vbCrLf
            endIt = True
            Exit Sub
        End If
    Next a
End Sub
