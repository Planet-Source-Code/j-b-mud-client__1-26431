VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ListBox lstServer 
      Height          =   2790
      ItemData        =   "frmConnect.frx":0000
      Left            =   120
      List            =   "frmConnect.frx":0002
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "IP/Address:"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Add
    If txtName.Text = "" Or txtIP.Text = "" Or txtPort.Text = "" Then
        MsgBox "You forgot to enter a name or IP # or port"
    Else
        Dim a As String
        a = lstServer.ListCount + 1
        KeySection = "Count"
        KeyKey = "ListCount"
        KeyValue = a
        SaveINI
        KeySection = "Names"
        KeyKey = "Name" & a
        KeyValue = txtName.Text
        SaveINI
        KeySection = "IPs"
        KeyKey = "IP" & a
        KeyValue = txtIP.Text
        SaveINI
        KeySection = "Ports"
        KeyKey = "Port" & a
        KeyValue = txtPort.Text
        SaveINI
        lstServer.AddItem txtName.Text
        txtName.Text = ""
        txtIP.Text = ""
        txtPort.Text = ""
    End If
End Sub

Private Sub Command2_Click() 'Remove
    KeySection = "Count"
    KeyKey = "ListCount"
    LoadINI
    Dim ListCount As Integer
    ListCount = KeyValue
    KeySection = "Names"
    KeyKey = "Name" & ListCount
    LoadINI
    Dim TempName As String
    TempName = KeyValue
    KeySection = "IPs"
    KeyKey = "IP" & ListCount
    LoadINI
    Dim TempIP As String
    TempIP = KeyValue
    KeySection = "Ports"
    KeyKey = "Port" & ListCount
    LoadINI
    Dim TempPort As String
    TempPort = KeyValue
    a = lstServer.ListIndex
    a = a + 1
    ReplaceServer Val(a), TempName, TempIP, TempPort
    ReplaceServer ListCount, vbNullString, vbNullString, vbNullString
    ListCount = ListCount - 1
    KeySection = "Count"
    KeyKey = "ListCount"
    KeyValue = ListCount
    SaveINI
    RefreshSList
End Sub

Private Sub Command3_Click() 'Replace
    a = lstServer.ListIndex
    a = a + 1
    ReplaceServer Val(a), txtName.Text, txtIP.Text, txtPort.Text
    RefreshSList
End Sub

Private Sub Command4_Click() 'Connect
    'doesn't work yet
    If InStr(1, KeyValue, "True") Then
        frmOptions.chkSize.Caption = "Save Size On"
        frmOptions.chkSize.Value = 1
        KeySection = "Size"
        KeyKey = "Height"
        LoadINI
        frmMain.Height = Val(KeyValue)
        KeySection = "Size"
        KeyKey = "Width"
        LoadINI
        frmMain.Width = Val(KeyValue)
        KeySection = "Size"
        KeyKey = "ScaleHeight"
        LoadINI
        frmMain.ScaleHeight = Val(KeyValue)
        KeySection = "Size"
        KeyKey = "ScaleWidth"
        LoadINI
        frmMain.ScaleWidth = Val(KeyValue)
    ElseIf InStr(1, KeyValue, "False") Then
        frmOptions.chkSize.Caption = "Save Size Off"
        frmOptions.chkSize.Value = 0
    End If
    'connect to the mud and show main form
    frmMain.Winsock1.RemoteHost = txtIP.Text
    frmMain.Winsock1.RemotePort = txtPort.Text
    frmMain.Winsock1.Connect
    frmMain.Caption = txtName.Text & " ::Connecting ..."
    frmMain.Show
    frmMain.mnuFileConnectToggle.Caption = "&Disconnect"
    Me.Hide
End Sub

Private Sub LoadINI()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\ServerList.ini" 'Declare your ini file !
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

Private Sub SaveINI()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\ServerList.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub

Private Sub Form_Load()
    RefreshSList
End Sub

Private Sub RefreshSList()
    On Error GoTo LoadError
    KeySection = "Count"
    KeyKey = "ListCount"
    LoadINI
    Dim ListCount As String
    ListCount = KeyValue
    Dim b As Integer
    Dim c As Integer
    Dim d As String
    c = Val(ListCount)
    lstServer.Clear
    For b = 1 To c
        d = b
        KeySection = "Names"
        KeyKey = "Name" + d
        LoadINI
        lstServer.AddItem KeyValue
    Next b
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            MsgBox "Your list is empty."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lstServer_Click()
    Dim Index As String
    Index = lstServer.ListIndex + 1
    KeySection = "Names"
    KeyKey = "Name" & Index
    LoadINI
    txtName.Text = KeyValue
    KeySection = "IPs"
    KeyKey = "IP" & Index
    LoadINI
    txtIP.Text = KeyValue
    KeySection = "Ports"
    KeyKey = "Port" & Index
    LoadINI
    txtPort.Text = KeyValue
End Sub

Private Sub ReplaceServer(Count As Integer, Name As String, IP As String, Port As String)
    KeySection = "Names"
    KeyKey = "Name" & Count
    KeyValue = Name
    SaveINI
    KeySection = "IPs"
    KeyKey = "IP" & Count
    KeyValue = IP
    SaveINI
    KeySection = "Ports"
    KeyKey = "Port" & Count
    KeyValue = Port
    SaveINI
End Sub

Private Sub lstServer_DblClick()
    lstServer_Click
    Command4_Click
End Sub
