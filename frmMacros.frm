VERSION 5.00
Begin VB.Form frmMacros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macros"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   300
      Left            =   2520
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "F12:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "F11:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "F10:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2090
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "F9:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1730
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "F8:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1370
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "F7:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1010
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "F6:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   650
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "F5:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   290
      Width           =   375
   End
End
Attribute VB_Name = "frmMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Done
    If Text1.Text = "" Then
        Text1.Text = " "
    End If
    If Text2.Text = "" Then
        Text2.Text = " "
    End If
    If Text2.Text = "" Then
        Text2.Text = " "
    End If
    If Text3.Text = "" Then
        Text3.Text = " "
    End If
    If Text4.Text = "" Then
        Text4.Text = " "
    End If
    If Text5.Text = "" Then
        Text5.Text = " "
    End If
    If Text6.Text = "" Then
        Text6.Text = " "
    End If
    If Text7Text = "" Then
        Text7.Text = " "
    End If
    If Text8.Text = "" Then
        Text8.Text = " "
    End If
    SaveINI "Settings", "Macros", "F5", Text1.Text
    SaveINI "Settings", "Macros", "F6", Text2.Text
    SaveINI "Settings", "Macros", "F7", Text3.Text
    SaveINI "Settings", "Macros", "F8", Text4.Text
    SaveINI "Settings", "Macros", "F9", Text5.Text
    SaveINI "Settings", "Macros", "F10", Text6.Text
    SaveINI "Settings", "Macros", "F11", Text7.Text
    SaveINI "Settings", "Macros", "F12", Text8.Text
    Unload Me
End Sub

Private Sub Command2_Click() 'Cancel
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo LoadError
    LoadINI "Settings", "Macros", "F5"
    Text1.Text = KeyValue
    LoadINI "Settings", "Macros", "F6"
    Text2.Text = KeyValue
    LoadINI "Settings", "Macros", "F7"
    Text3.Text = KeyValue
    LoadINI "Settings", "Macros", "F8"
    Text4.Text = KeyValue
    LoadINI "Settings", "Macros", "F9"
    Text5.Text = KeyValue
    LoadINI "Settings", "Macros", "F10"
    Text6.Text = KeyValue
    LoadINI "Settings", "Macros", "F11"
    Text7.Text = KeyValue
    LoadINI "Settings", "Macros", "F12"
    Text8.Text = KeyValue
LoadError:
    ErrorNumber = Err.Number
    Select Case ErrorNumber
        Case 13
            MsgBox "Your macros are empty."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
End Sub
