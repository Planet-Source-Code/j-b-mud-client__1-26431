VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSize 
      Caption         =   "Save Size Off"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Value           =   2  'Grayed
      Width           =   1575
   End
   Begin VB.CheckBox chkAlias 
      Caption         =   "Alias On"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkEcho 
      Caption         =   "Echo Off (repeats to you what you enter)"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAlias_Click()
    If chkAlias.Value = 0 Then
        chkAlias.Caption = "Alias Off"
        KeySection = "Options"
        KeyKey = "Alias"
        KeyValue = "False"
        SaveINI
    ElseIf chkAlias.Value = 1 Then
        chkAlias.Caption = "Alias On"
        KeySection = "Options"
        KeyKey = "Alias"
        KeyValue = "True"
        SaveINI
    End If
End Sub

Private Sub chkEcho_Click()
    If chkEcho.Value = 0 Then
        chkEcho.Caption = "Echo Off (repeats to you what you enter)"
        KeySection = "Options"
        KeyKey = "Echo"
        KeyValue = "False"
        SaveINI
    ElseIf chkEcho.Value = 1 Then
        chkEcho.Caption = "Echo On (repeats to you what you enter)"
        KeySection = "Options"
        KeyKey = "Echo"
        KeyValue = "True"
        SaveINI
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
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub

Private Sub chkSize_Click()
    If chkSize.Value = 0 Then
        chkSize.Caption = "Save Size Off"
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "False"
        SaveINI
    ElseIf chkSize.Value = 1 Then
        chkSize.Caption = "Save Size On"
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "True"
        SaveINI
        KeySection = "Size"
        KeyKey = "Height"
        KeyValue = frmMain.Height
        SaveINI
        KeySection = "Size"
        KeyKey = "Width"
        KeyValue = frmMain.Width
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleHeight"
        KeyValue = frmMain.ScaleHeight
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleWidth"
        KeyValue = frmMain.ScaleWidth
        SaveINI
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo LoadError
    KeySection = "Options"
    KeyKey = "Alias"
    LoadINI
    If InStr(1, KeyValue, "True") Then
        chkAlias.Caption = "Alias On"
        chkAlias.Value = 1
    ElseIf InStr(1, KeyValue, "False") Then
        chkAlias.Caption = "Alias Off"
        chkAlias.Value = 0
    End If
    KeySection = "Options"
    KeyKey = "Echo"
    LoadINI
    If InStr(1, KeyValue, "True") Then
        chkEcho.Caption = "Echo On (repeats to you what you enter)"
        chkEcho.Value = 1
    ElseIf InStr(1, KeyValue, "False") Then
        chkEcho.Caption = "Echo Off (repeats to you what you enter)"
        chkEcho.Value = 0
    End If
    KeySection = "Options"
    KeyKey = "Size"
    LoadINI
    If InStr(1, KeyValue, "True") Then
        chkSize.Caption = "Save Size On"
        chkSize.Value = 1
    ElseIf InStr(1, KeyValue, "False") Then
        chkSize.Caption = "Save Size Off"
        chkSize.Value = 0
    End If
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            'do nothing
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If chkAlias.Value = 0 Then
        KeySection = "Options"
        KeyKey = "Alias"
        KeyValue = "False"
        SaveINI
    ElseIf chkAlias.Value = 1 Then
        KeySection = "Options"
        KeyKey = "Alias"
        KeyValue = "True"
        SaveINI
    End If
    If chkEcho.Value = 0 Then
        KeySection = "Options"
        KeyKey = "Echo"
        KeyValue = "False"
        SaveINI
    ElseIf chkEcho.Value = 1 Then
        KeySection = "Options"
        KeyKey = "Echo"
        KeyValue = "True"
        SaveINI
    End If
    If chkSize.Value = 0 Then
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "False"
        SaveINI
    ElseIf chkSize.Value = 1 Then
        KeySection = "Options"
        KeyKey = "Size"
        KeyValue = "True"
        SaveINI
        KeySection = "Size"
        KeyKey = "Height"
        KeyValue = frmMain.Height
        SaveINI
        KeySection = "Size"
        KeyKey = "Width"
        KeyValue = frmMain.Width
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleHeight"
        KeyValue = frmMain.ScaleHeight
        SaveINI
        KeySection = "Size"
        KeyKey = "ScaleWidth"
        KeyValue = frmMain.ScaleWidth
        SaveINI
    End If
    Me.Hide
End Sub
