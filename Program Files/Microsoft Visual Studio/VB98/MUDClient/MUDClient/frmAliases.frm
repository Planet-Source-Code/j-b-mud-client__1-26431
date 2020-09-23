VERSION 5.00
Begin VB.Form frmAliases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aliases"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstAliases 
      Height          =   2790
      ItemData        =   "frmAliases.frx":0000
      Left            =   120
      List            =   "frmAliases.frx":0002
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtAlias 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Command:"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Alias:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAliases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click() 'Add
    If txtAlias.Text = "" Or txtCommand.Text = "" Then
        MsgBox "You forgot to enter an Alias or Command"
    Else
        Dim a As String
        a = lstAliases.ListCount + 1
        KeySection = "Count"
        KeyKey = "AliasCount"
        KeyValue = a
        SaveINI
        KeySection = "Aliases"
        KeyKey = "Alias" & a
        KeyValue = txtAlias.Text
        SaveINI
        KeySection = "Commands"
        KeyKey = "Command" & a
        KeyValue = txtCommand.Text
        SaveINI
        lstAliases.AddItem txtAlias.Text
        txtAlias.Text = ""
        txtCommand.Text = ""
    End If
End Sub

Private Sub cmdRemove_Click() 'Remove
    KeySection = "Count"
    KeyKey = "AliasCount"
    LoadINI
    Dim AliasCount As Integer
    AliasCount = KeyValue
    KeySection = "Aliases"
    KeyKey = "Alias" & AliasCount
    LoadINI
    Dim TempAlias As String
    TempAlias = KeyValue
    KeySection = "Commands"
    KeyKey = "Command" & AliasCount
    LoadINI
    Dim TempCommand As String
    TempCommand = KeyValue
    a = lstAliases.ListIndex
    a = a + 1
    ReplaceAlias Val(a), TempAlias, TempCommand
    ReplaceAlias AliasCount, vbNullString, vbNullString
    AliasCount = AliasCount - 1
    KeySection = "Count"
    KeyKey = "AliasCount"
    KeyValue = AliasCount
    SaveINI
    RefreshAList
End Sub

Private Sub cmdReplace_Click() 'Replace
    a = lstAliases.ListIndex
    a = a + 1
    ReplaceAlias Val(a), txtAlias.Text, txtCommand.Text
    RefreshAList
End Sub

Private Sub cmdClose_Click() 'Connect
    Me.Hide
End Sub

Private Sub LoadINI()

Dim lngResult As Long
Dim strFileAlias
Dim strResult As String * 50
strFileAlias = App.Path & "\AliasList.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileAlias, strResult, Len(strResult), _
strFileAlias)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub

Private Sub SaveINI()

Dim lngResult As Long
Dim strFileAlias
strFileAlias = App.Path & "\AliasList.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileAlias)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub

Private Sub Form_Load()
    RefreshAList
End Sub

Private Sub RefreshAList()
    On Error GoTo LoadError
    KeySection = "Count"
    KeyKey = "AliasCount"
    LoadINI
    Dim AliasCount As String
    AliasCount = KeyValue
    Dim b As Integer
    Dim c As Integer
    Dim d As String
    c = Val(AliasCount)
    lstAliases.Clear
    For b = 1 To c
        d = b
        KeySection = "Aliases"
        KeyKey = "Alias" + d
        LoadINI
        lstAliases.AddItem KeyValue
    Next b
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            MsgBox "Your list is empty."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
End Sub

Private Sub lstAliases_Click()
    Dim Index As String
    Index = lstAliases.ListIndex + 1
    KeySection = "Aliases"
    KeyKey = "Alias" & Index
    LoadINI
    txtAlias.Text = KeyValue
    KeySection = "Commands"
    KeyKey = "Command" & Index
    LoadINI
    txtCommand.Text = KeyValue
End Sub

Private Sub ReplaceAlias(Count As Integer, Alias As String, Command As String)
    KeySection = "Aliases"
    KeyKey = "Alias" & Count
    KeyValue = Alias
    SaveINI
    KeySection = "Commands"
    KeyKey = "Command" & Count
    KeyValue = Command
    SaveINI
End Sub
