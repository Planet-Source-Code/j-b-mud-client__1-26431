Attribute VB_Name = "Module1"
Declare Function WritePrivateProfileString _
Lib "KERNEL32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lsString As Any, _
ByVal lplFilename As String) As Long

Declare Function GetPrivateProfileString Lib _
"KERNEL32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal _
lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As _
String) As Long

Public KeySection As String
Public KeyKey As String
Public KeyValue As String
Public endIt As Boolean
Public MainHeight As Long
Public MainWidth As Long
Public MainScaleHeight As Long
Public MainScaleWidth As Long
Public LastColor As String
Public LastColorCode As String

'\B/-----------------------------------Ansi colors-----------------------------------
Public BLACK As String ' = "[1m[30m"
Public RED As String ' = "[1m[31m"
Public GREEN As String ' = "[1m[32m"
Public YELLOW As String ' = "[1m[33m"
Public BLUE As String ' = "[1m[34m"
Public MAGENTA As String ' = "[1m[35m"
Public LIGHTBLUE As String ' = "[1m[36m"
Public WHITE As String ' = "[1m[37m"

Public bBLACK As String ' = "[0m[30m"
Public bRED As String ' = "[0m[31m"
Public bGREEN As String ' = "[0m[32m"
Public bYELLOW As String ' = "[0m[33m"
Public bBLUE As String ' = "[0m[34m"
Public bMAGENTA As String ' = "[0m[35m"
Public bLIGHTBLUE As String ' = "[0m[36m"
Public bWHITE As String ' = "[0m[37m"
'/E\-----------------------------------Ansi colors-----------------------------------

'\B/-----------------------------------Ansi colors2----------------------------------
Public Const BLACK2 As String = "[1m[30m"
Public Const RED2 As String = "[1m[31m"
Public Const GREEN2 As String = "[1m[32m"
Public Const YELLOW2 As String = "[1m[33m"
Public Const BLUE2 As String = "[1m[34m"
Public Const MAGENTA2 As String = "[1m[35m"
Public Const LIGHTBLUE2 As String = "[1m[36m"
Public Const WHITE2 As String = "[1m[37m"

Public Const bBLACK2 As String = "[0m[30m"
Public Const bRED2 As String = "[0m[31m"
Public Const bGREEN2 As String = "[0m[32m"
Public Const bYELLOW2 As String = "[0m[33m"
Public Const bBLUE2 As String = "[0m[34m"
Public Const bMAGENTA2 As String = "[0m[35m"
Public Const bLIGHTBLUE2 As String = "[0m[36m"
Public Const bWHITE2 As String = "[0m[37m"
'/E\-----------------------------------Ansi colors2----------------------------------

