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
      TextRTF         =   $"frmMain.frx":0000
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
Option Explicit

Dim NextColor As Boolean
Dim AColor As String

Private Sub CommandBar_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ErrorNumber As Integer
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
    Dim ErrorNumber As Integer
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
                If frmOptions.chkEcho.Value = 1 Then Display YELLOW2 & CommandBar.Text & vbCrLf & LastColorCode
                If frmOptions.chkEcho.Value = 0 Then Display vbCrLf & LastColorCode
                CommandBar.SelStart = 0
                CommandBar.SelLength = Len(CommandBar.Text)
                If endIt Then Exit Sub
            Else
                If frmOptions.chkEcho.Value = 1 Then Display YELLOW2 & CommandBar.Text & vbCrLf & LastColorCode
                If frmOptions.chkEcho.Value = 0 Then Display vbCrLf & LastColorCode
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
        tmpDisplay = WHITE2 & vbCrLf & "***Disconnected***" & vbCrLf & LastColorCode
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
    tmpDisplay = WHITE2 & vbCrLf & "***Disconnected***" & vbCrLf & LastColorCode
    Display tmpDisplay
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    mnuFileConnectToggle.Caption = "&Disconnect"
    Me.Caption = frmConnect.txtName.Text & " ::Connected"
    Dim tmpDisplay As String
    tmpDisplay = WHITE2 & vbCrLf & "Connected to " & Winsock1.RemoteHost & LastColorCode
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim UserCommand As String
    Dim WrkSpace As String
    Winsock1.GetData UserCommand
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
    Dim InTemp As Long
    Dim LastBLACK As Long
    Dim LastRED As Long
    Dim LastGREEN As Long
    Dim LastYELLOW As Long
    Dim LastBLUE As Long
    Dim LastMAGENTA As Long
    Dim LastLIGHTBLUE As Long
    Dim LastWHITE As Long
    Dim LastbBLACK As Long
    Dim LastbRED As Long
    Dim LastbGREEN As Long
    Dim LastbYELLOW As Long
    Dim LastbBLUE As Long
    Dim LastbMAGENTA As Long
    Dim LastbLIGHTBLUE As Long
    Dim LastbWHITE As Long
    'MsgBox WrkSpace
    'Clipboard.SetText WrkSpace
    WrkSpace2 = WrkSpace
    'Format the message so it can be fed into the display box's RTF
    WrkSpace = Replace(WrkSpace, "\", "\\")     '
    WrkSpace = Replace(WrkSpace, "{", "\{")     'these three replace stuff that could be harmful to the code
    WrkSpace = Replace(WrkSpace, "}", "\}")     '
    WrkSpace = Replace(WrkSpace, vbCr, vbCrLf)              'takes care of returns ... just not sure as to how well though
    'WrkSpace = Replace(WrkSpace, vbLg, vbCrLf)              '
    WrkSpace = Replace(WrkSpace, vbCrLf, vbCrLf & "\par ")  '
    WrkSpace = Replace(WrkSpace, "ÿû", " ")    'gets rid of tags for Name and Password entries at login
    WrkSpace = Replace(WrkSpace, "ÿü", vbCrLf) '
    'this is the begining of the long process of detecting and replacing colors as well as detecting the last color to appear in the message
    GoTo BeginBLACK
BeginBLACK:
    'find the color
    FindColor WrkSpace, 1, 0
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginRED
    Else
        'FindColor() found the color so now we define the color and continue
        BLACK = AColor
        GoTo ContinueBLACK
    End If
ContinueBLACK:
    'find the last occurance of this color in the message
    strTemp = BLACK
    InTemp = 1
    LastBLACK = 0
    Do While InStr(InTemp, WrkSpace, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastBLACK = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, BLACK, "\plain\f2\fs20\cf0 ")
    GoTo BeginRED
BeginRED:
    'find the color
    FindColor WrkSpace, 1, 1
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginGREEN
    Else
        'FindColor() found the color so now we define the color and continue
        RED = AColor
        GoTo ContinueRED
    End If
ContinueRED:
    'find the last occurance of this color in the message
    strTemp = RED
    InTemp = 1
    LastRED = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastRED = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, RED, "\plain\f2\fs20\cf2 ")
    GoTo BeginGREEN
BeginGREEN:
    'find the color
    FindColor WrkSpace, 1, 2
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginYELLOW
    Else
        'FindColor() found the color so now we define the color and continue
        GREEN = AColor
        GoTo ContinueGREEN
    End If
ContinueGREEN:
    'find the last occurance of this color in the message
    strTemp = GREEN
    InTemp = 1
    LastGREEN = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastGREEN = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, GREEN, "\plain\f2\fs20\cf3 ")
    GoTo BeginYELLOW
BeginYELLOW:
    'find the color
    FindColor WrkSpace, 1, 3
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginBLUE
    Else
        'FindColor() found the color so now we define the color and continue
        YELLOW = AColor
        GoTo ContinueYELLOW
    End If
ContinueYELLOW:
    'find the last occurance of this color in the message
    strTemp = YELLOW
    InTemp = 1
    LastYELLOW = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastYELLOW = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, YELLOW, "\plain\f2\fs20\cf4 ")
    GoTo BeginBLUE
BeginBLUE:
    'find the color
    FindColor WrkSpace, 1, 4
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginMAGENTA
    Else
        'FindColor() found the color so now we define the color and continue
        BLUE = AColor
        GoTo ContinueBLUE
    End If
ContinueBLUE:
    'find the last occurance of this color in the message
    strTemp = BLUE
    InTemp = 1
    LastBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, BLUE, "\plain\f2\fs20\cf1 ")
    GoTo BeginMAGENTA
BeginMAGENTA:
    'find the color
    FindColor WrkSpace, 1, 5
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginLIGHTBLUE
    Else
        'FindColor() found the color so now we define the color and continue
        MAGENTA = AColor
        GoTo ContinueMAGENTA
    End If
ContinueMAGENTA:
    'find the last occurance of this color in the message
    strTemp = MAGENTA
    InTemp = 1
    LastMAGENTA = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastMAGENTA = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, MAGENTA, "\plain\f2\fs20\cf7 ")
    GoTo BeginLIGHTBLUE
BeginLIGHTBLUE:
    'find the color
    FindColor WrkSpace, 1, 6
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginWHITE
    Else
        'FindColor() found the color so now we define the color and continue
        LIGHTBLUE = AColor
        GoTo ContinueLIGHTBLUE
    End If
ContinueLIGHTBLUE:
    'find the last occurance of this color in the message
    strTemp = LIGHTBLUE
    InTemp = 1
    LastLIGHTBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastLIGHTBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, LIGHTBLUE, "\plain\f2\fs20\cf5 ")
    GoTo BeginWHITE
BeginWHITE:
    'find the color
    FindColor WrkSpace, 1, 7
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbBLACK
    Else
        'FindColor() found the color so now we define the color and continue
        WHITE = AColor
        GoTo ContinueWHITE
    End If
ContinueWHITE:
    'find the last occurance of this color in the message
    strTemp = WHITE
    InTemp = 1
    LastbWHITE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbWHITE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, WHITE, "\plain\f2\fs20\cf6 ")
    GoTo BeginbBLACK
BeginbBLACK:
    'find the color
    FindColor WrkSpace, 0, 0
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbRED
    Else
        'FindColor() found the color so now we define the color and continue
        bBLACK = AColor
        GoTo ContinuebBLACK
    End If
ContinuebBLACK:
    'find the last occurance of this color in the message
    strTemp = bBLACK
    InTemp = 1
    LastbBLACK = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbBLACK = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bBLACK, "\plain\f2\fs20\cf0 ")
    GoTo BeginbRED
BeginbRED:
    'find the color
    FindColor WrkSpace, 0, 1
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbGREEN
    Else
        'FindColor() found the color so now we define the color and continue
        bRED = AColor
        GoTo ContinuebRED
    End If
ContinuebRED:
    'find the last occurance of this color in the message
    strTemp = bRED
    InTemp = 1
    LastbRED = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbRED = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bRED, "\plain\f2\fs20\cf10 ")
    GoTo BeginbGREEN
BeginbGREEN:
    'find the color
    FindColor WrkSpace, 0, 2
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbYELLOW
    Else
        'FindColor() found the color so now we define the color and continue
        bGREEN = AColor
        GoTo ContinuebGREEN
    End If
ContinuebGREEN:
    'find the last occurance of this color in the message
    strTemp = bGREEN
    InTemp = 1
    LastbGREEN = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbGREEN = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bGREEN, "\plain\f2\fs20\cf9 ")
    GoTo BeginbYELLOW
BeginbYELLOW:
    'find the color
    FindColor WrkSpace, 0, 3
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbBLUE
    Else
        'FindColor() found the color so now we define the color and continue
        bYELLOW = AColor
        GoTo ContinuebYELLOW
    End If
ContinuebYELLOW:
    'find the last occurance of this color in the message
    strTemp = bYELLOW
    InTemp = 1
    LastbYELLOW = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbYELLOW = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bYELLOW, "\plain\f2\fs20\cf11 ")
    GoTo BeginbBLUE
BeginbBLUE:
    'find the color
    FindColor WrkSpace, 0, 4
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbMAGENTA
    Else
        'FindColor() found the color so now we define the color and continue
        bBLUE = AColor
        GoTo ContinuebBLUE
    End If
ContinuebBLUE:
    'find the last occurance of this color in the message
    strTemp = bBLUE
    InTemp = 1
    LastbBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bBLUE, "\plain\f2\fs20\cf12 ")
    GoTo BeginbMAGENTA
BeginbMAGENTA:
    'find the color
    FindColor WrkSpace, 0, 5
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbLIGHTBLUE
    Else
        'FindColor() found the color so now we define the color and continue
        bMAGENTA = AColor
        GoTo ContinuebMAGENTA
    End If
ContinuebMAGENTA:
    'find the last occurance of this color in the message
    strTemp = bMAGENTA
    InTemp = 1
    LastbMAGENTA = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbMAGENTA = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bMAGENTA, "\plain\f2\fs20\cf14 ")
    GoTo BeginbLIGHTBLUE
BeginbLIGHTBLUE:
    'find the color
    FindColor WrkSpace, 0, 6
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo BeginbWHITE
    Else
        'FindColor() found the color so now we define the color and continue
        bLIGHTBLUE = AColor
        GoTo ContinuebLIGHTBLUE
    End If
ContinuebLIGHTBLUE:
    'find the last occurance of this color in the message
    strTemp = bLIGHTBLUE
    InTemp = 1
    LastbLIGHTBLUE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbLIGHTBLUE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bLIGHTBLUE, "\plain\f2\fs20\cf15 ")
    GoTo BeginbWHITE
BeginbWHITE:
    'find the color
    FindColor WrkSpace, 0, 7
    If NextColor Then
        'FindColor() could not find this color so we go to the next one
        GoTo EndbWHITE
    Else
        'FindColor() found the color so now we define the color and continue
        bWHITE = AColor
        GoTo ContinuebWHITE
    End If
ContinuebWHITE:
    'find the last occurance of this color in the message
    strTemp = bWHITE
    InTemp = 1
    LastbWHITE = 0
    Do While InStr(InTemp, WrkSpace2, strTemp) > 0 & InStr(InTemp, WrkSpace2, strTemp) < Len(WrkSpace2)
        If InStr(InTemp, WrkSpace, strTemp) = 0 Then Exit Do
        LastbWHITE = InTemp
        InTemp = InStr(InTemp + 1, WrkSpace2, strTemp)
        If InTemp = 0 Then Exit Do
    Loop
    'replace all occurances of this color with code the text box can understand
    WrkSpace = Replace(WrkSpace, bWHITE, "\plain\f2\fs20\cf8 ")
    GoTo EndbWHITE
EndbWHITE:
    'compare to see which color is the last one
    If LastBLACK > LastRED & LastBLACK > LastGREEN & LastBLACK > LastYELLOW & LastBLACK > LastBLUE & LastBLACK > LastMAGENTA & LastBLACK > LastLIGHTBLUE & LastBLACK > LastWHITE & LastBLACK > LastbBLACK & LastBLACK > LastbRED & LastBLACK > LastbGREEN & LastBLACK > LastbYELLOW & LastBLACK > LastbBLUE & LastBLACK > LastbMAGENTA & LastBLACK > LastbLIGHTBLUE & LastBLACK > LastbWHITE Then
        LastColor = "0"
        LastColorCode = BLACK
    ElseIf LastRED > LastBLACK & LastRED > LastGREEN & LastRED > LastYELLOW & LastRED > LastBLUE & LastRED > LastMAGENTA & LastRED > LastLIGHTBLUE & LastRED > LastWHITE & LastRED > LastbBLACK & LastRED > LastbRED & LastRED > LastbGREEN & LastRED > LastbYELLOW & LastRED > LastbBLUE & LastRED > LastbMAGENTA & LastRED > LastbLIGHTBLUE & LastRED > LastbWHITE Then
        LastColor = "2"
        LastColorCode = RED
    ElseIf LastGREEN > LastBLACK & LastGREEN > LastRED & LastGREEN > LastYELLOW & LastGREEN > LastBLUE & LastGREEN > LastMAGENTA & LastGREEN > LastLIGHTBLUE & LastGREEN > LastWHITE & LastGREEN > LastbBLACK & LastGREEN > LastbRED & LastGREEN > LastbGREEN & LastGREEN > LastbYELLOW & LastGREEN > LastbBLUE & LastGREEN > LastbMAGENTA & LastGREEN > LastbLIGHTBLUE & LastGREEN > LastbWHITE Then
        LastColor = "3"
        LastColorCode = GREEN
    ElseIf LastYELLOW > LastBLACK & LastYELLOW > LastRED & LastYELLOW > LastGREEN & LastYELLOW > LastBLUE & LastYELLOW > LastMAGENTA & LastYELLOW > LastLIGHTBLUE & LastYELLOW > LastWHITE & LastYELLOW > LastbBLACK & LastYELLOW > LastbRED & LastYELLOW > LastbGREEN & LastYELLOW > LastbYELLOW & LastYELLOW > LastbBLUE & LastYELLOW > LastbMAGENTA & LastYELLOW > LastbLIGHTBLUE & LastYELLOW > LastbWHITE Then
        LastColor = "4"
        LastColorCode = YELLOW
    ElseIf LastBLUE > LastBLACK & LastBLUE > LastRED & LastBLUE > LastGREEN & LastBLUE > LastYELLOW & LastBLUE > LastMAGENTA & LastBLUE > LastLIGHTBLUE & LastBLUE > LastWHITE & LastBLUE > LastbBLACK & LastBLUE > LastbRED & LastBLUE > LastbGREEN & LastBLUE > LastbYELLOW & LastBLUE > LastbBLUE & LastBLUE > LastbMAGENTA & LastBLUE > LastbLIGHTBLUE & LastBLUE > LastbWHITE Then
        LastColor = "1"
        LastColorCode = BLUE
    ElseIf LastMAGENTA > LastBLACK & LastMAGENTA > LastRED & LastMAGENTA > LastGREEN & LastMAGENTA > LastYELLOW & LastMAGENTA > LastBLUE & LastMAGENTA > LastLIGHTBLUE & LastMAGENTA > LastWHITE & LastMAGENTA > LastbBLACK & LastMAGENTA > LastbRED & LastMAGENTA > LastbGREEN & LastMAGENTA > LastbYELLOW & LastMAGENTA > LastbBLUE & LastMAGENTA > LastbMAGENTA & LastMAGENTA > LastbLIGHTBLUE & LastMAGENTA > LastbWHITE Then
        LastColor = "7"
        LastColorCode = MAGENTA
    ElseIf LastLIGHTBLUE > LastBLACK & LastLIGHTBLUE > LastRED & LastLIGHTBLUE > LastGREEN & LastLIGHTBLUE > LastYELLOW & LastLIGHTBLUE > LastBLUE & LastLIGHTBLUE > LastMAGENTA & LastLIGHTBLUE > LastWHITE & LastLIGHTBLUE > LastbBLACK & LastLIGHTBLUE > LastbRED & LastLIGHTBLUE > LastbGREEN & LastLIGHTBLUE > LastbYELLOW & LastLIGHTBLUE > LastbBLUE & LastLIGHTBLUE > LastbMAGENTA & LastLIGHTBLUE > LastbLIGHTBLUE & LastLIGHTBLUE > LastbWHITE Then
        LastColor = "5"
        LastColorCode = LIGHTBLUE
    ElseIf LastWHITE > LastBLACK & LastWHITE > LastRED & LastWHITE > LastGREEN & LastWHITE > LastYELLOW & LastWHITE > LastBLUE & LastWHITE > LastMAGENTA & LastWHITE > LastLIGHTBLUE & LastWHITE > LastbBLACK & LastWHITE > LastbRED & LastWHITE > LastbGREEN & LastWHITE > LastbYELLOW & LastWHITE > LastbBLUE & LastWHITE > LastbMAGENTA & LastWHITE > LastbLIGHTBLUE & LastWHITE > LastbWHITE Then
        LastColor = "6"
        LastColorCode = WHITE
    ElseIf LastbBLACK > LastRED & LastbBLACK > LastGREEN & LastbBLACK > LastYELLOW & LastbBLACK > LastBLUE & LastbBLACK > LastMAGENTA & LastbBLACK > LastLIGHTBLUE & LastbBLACK > LastWHITE & LastbBLACK > LastBLACK & LastbBLACK > LastbRED & LastbBLACK > LastbGREEN & LastbBLACK > LastbYELLOW & LastbBLACK > LastbBLUE & LastbBLACK > LastbMAGENTA & LastbBLACK > LastbLIGHTBLUE & LastbBLACK > LastbWHITE Then
        LastColor = "0"
        LastColorCode = bBLACK
    ElseIf LastbRED > LastBLACK & LastbRED > LastGREEN & LastbRED > LastYELLOW & LastbRED > LastBLUE & LastbRED > LastMAGENTA & LastbRED > LastLIGHTBLUE & LastbRED > LastWHITE & LastbRED > LastbBLACK & LastbRED > LastRED & LastbRED > LastbGREEN & LastbRED > LastbYELLOW & LastbRED > LastbBLUE & LastbRED > LastbMAGENTA & LastbRED > LastbLIGHTBLUE & LastbRED > LastbWHITE Then
        LastColor = "10"
        LastColorCode = bRED
    ElseIf LastbGREEN > LastBLACK & LastbGREEN > LastRED & LastbGREEN > LastYELLOW & LastbGREEN > LastBLUE & LastbGREEN > LastMAGENTA & LastbGREEN > LastLIGHTBLUE & LastbGREEN > LastWHITE & LastbGREEN > LastbBLACK & LastbGREEN > LastbRED & LastbGREEN > LastGREEN & LastbGREEN > LastbYELLOW & LastbGREEN > LastbBLUE & LastbGREEN > LastbMAGENTA & LastbGREEN > LastbLIGHTBLUE & LastbGREEN > LastbWHITE Then
        LastColor = "9"
        LastColorCode = bGREEN
    ElseIf LastbYELLOW > LastBLACK & LastbYELLOW > LastRED & LastbYELLOW > LastGREEN & LastbYELLOW > LastBLUE & LastbYELLOW > LastMAGENTA & LastbYELLOW > LastLIGHTBLUE & LastbYELLOW > LastWHITE & LastbYELLOW > LastbBLACK & LastbYELLOW > LastbRED & LastbYELLOW > LastbGREEN & LastbYELLOW > LastYELLOW & LastbYELLOW > LastbBLUE & LastbYELLOW > LastbMAGENTA & LastbYELLOW > LastbLIGHTBLUE & LastbYELLOW > LastbWHITE Then
        LastColor = "11"
        LastColorCode = bYELLOW
    ElseIf LastbBLUE > LastBLACK & LastbBLUE > LastRED & LastbBLUE > LastGREEN & LastbBLUE > LastYELLOW & LastbBLUE > LastMAGENTA & LastbBLUE > LastLIGHTBLUE & LastbBLUE > LastWHITE & LastbBLUE > LastbBLACK & LastbBLUE > LastbRED & LastbBLUE > LastbGREEN & LastbBLUE > LastbYELLOW & LastbBLUE > LastBLUE & LastbBLUE > LastbMAGENTA & LastbBLUE > LastbLIGHTBLUE & LastbBLUE > LastbWHITE Then
        LastColor = "12"
        LastColorCode = bBLUE
    ElseIf LastbMAGENTA > LastBLACK & LastbMAGENTA > LastRED & LastbMAGENTA > LastGREEN & LastbMAGENTA > LastYELLOW & LastbMAGENTA > LastBLUE & LastbMAGENTA > LastLIGHTBLUE & LastbMAGENTA > LastWHITE & LastbMAGENTA > LastbBLACK & LastbMAGENTA > LastbRED & LastbMAGENTA > LastbGREEN & LastbMAGENTA > LastbYELLOW & LastbMAGENTA > LastbBLUE & LastbMAGENTA > LastbMAGENTA & LastbMAGENTA > LastbLIGHTBLUE & LastbMAGENTA > LastbWHITE Then
        LastColor = "14"
        LastColorCode = bMAGENTA
    ElseIf LastbLIGHTBLUE > LastBLACK & LastbLIGHTBLUE > LastRED & LastbLIGHTBLUE > LastGREEN & LastbLIGHTBLUE > LastYELLOW & LastbLIGHTBLUE > LastBLUE & LastbLIGHTBLUE > LastMAGENTA & LastbLIGHTBLUE > LastWHITE & LastbLIGHTBLUE > LastbBLACK & LastbLIGHTBLUE > LastbRED & LastbLIGHTBLUE > LastbGREEN & LastbLIGHTBLUE > LastbYELLOW & LastbLIGHTBLUE > LastbBLUE & LastbLIGHTBLUE > LastbMAGENTA & LastbLIGHTBLUE > LastLIGHTBLUE & LastbLIGHTBLUE > LastbWHITE Then
        LastColor = "15"
        LastColorCode = bLIGHTBLUE
    ElseIf LastbWHITE > LastBLACK & LastbWHITE > LastRED & LastbWHITE > LastGREEN & LastbWHITE > LastYELLOW & LastbWHITE > LastBLUE & LastbWHITE > LastMAGENTA & LastbWHITE > LastLIGHTBLUE & LastbWHITE > LastbBLACK & LastbWHITE > LastbRED & LastbWHITE > LastbGREEN & LastbWHITE > LastbYELLOW & LastbWHITE > LastbBLUE & LastbWHITE > LastbMAGENTA & LastbWHITE > LastbLIGHTBLUE & LastbWHITE > LastWHITE Then
        LastColor = "8"
        LastColorCode = bWHITE
    End If
    'adding stuff to the beginning so it will work right (defines the font and different colors and other neccesary code)
    If LastColor = "" Then LastColor = "3"
    WrkSpace = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fmodern OCR A Extended;}}" & vbCrLf & "{\colortbl\red127\green127\blue127;\red0\green0\blue255;\red255\green0\blue0;\red0\green255\blue0;\red255\green255\blue0;\red128\green255\blue255;\red255\green255\blue255;\red255\green0\blue175;\red200\green200\blue200;\red0\green125\blue0;\red150\green0\blue0;\red125\green125\blue0;\red0\green0\blue125;\red0\green0\blue0;\red151\green0\blue102;\red0\green150\blue150;}" & vbCrLf & "\deflang1033\pard\plain\f2\fs20\cf" & LastColor & " " & WrkSpace
    'adding stuff to the end so it'll work right
    WrkSpace = WrkSpace & vbCrLf & "\plain\f2\fs20\par }"
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
    Dim a As Integer
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

Private Sub FindColor(WrkSpace As String, ClrType As Integer, Color As Integer)
    Dim TempBegin As Integer
    Dim TempBegin2 As Integer
    Dim TempStrLength As Integer
    Dim TempColor As String
    TempBegin = 0
    TempBegin2 = 0
    If InStr(1, WrkSpace, "3" & Color & "m") Then
        Do While InStr(TempBegin2 + 1, WrkSpace, "3" & Color & "m") > 0
            TempBegin2 = InStr(TempBegin2 + 1, WrkSpace, "3" & Color & "m")
            TempBegin = 0
            Do While InStr(TempBegin + 1, WrkSpace, "[" & ClrType) > 0
                TempBegin = InStr(TempBegin + 1, WrkSpace, "[" & ClrType)
                If TempBegin > TempBegin2 Then Exit Do
                TempColor = Mid(WrkSpace, TempBegin, (TempBegin2 - TempBegin) + 3)
                TempStrLength = Len(TempColor)
                If TempStrLength < 11 Then
                If Mid(TempColor, 4, 1) = ";" Or Mid(TempColor, 4, 1) = "" Or Mid(TempColor, 5, 1) = ";" Or Mid(TempColor, 5, 1) = "" Then
                    'the color has been found now go to finding the last position of the color and replacing the color in the message
                    AColor = TempColor
                    NextColor = False
                    Exit Sub
                End If
                End If
            Loop
        Loop
        Do While InStr(TempBegin2 + 1, WrkSpace, "[" & ClrType) > 0
            TempBegin2 = InStr(TempBegin2 + 1, WrkSpace, "[" & ClrType)
            TempBegin = 0
            Do While InStr(TempBegin + 1, WrkSpace, "3" & Color & "m") > 0
                TempBegin = InStr(TempBegin + 1, WrkSpace, "3" & Color & "m")
                If TempBegin > TempBegin2 Then Exit Do
                TempColor = Mid(WrkSpace, TempBegin, (TempBegin2 - TempBegin) + 3)
                TempStrLength = Len(TempColor)
                If TempStrLength < 11 Then
                If Mid(TempColor, 4, 1) = ";" Or Mid(TempColor, 4, 1) = "" Then
                    'the color has been found now go to finding the last position of the color and replacing the color in the message
                    AColor = TempColor
                    NextColor = False
                    Exit Sub
                End If
                End If
            Loop
        Loop
        NextColor = True
    Else
        'the color was not found go to next color
        NextColor = True
    End If
End Sub
