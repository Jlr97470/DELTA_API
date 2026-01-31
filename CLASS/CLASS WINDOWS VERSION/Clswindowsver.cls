VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Type OsVersionInfo
'    dwVersionInfoSize As Long
'    dwMajorVersion As Long
'    dwMinorVersion As Long
'    dwBuildNumber As Long
'    dwPlatform As Long
'    szCSDVersion As String * 128
'End Type
'Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpStruct As OsVersionInfo)
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'Private OsVers As OsVersionInfo
'Private winOS&
'Private Const Win95 = 1&
'Private Const WinNT = 2&
'
'Sub GetVersion32()
'    OsVers.dwVersionInfoSize = 148&
'    tempLong = GetVersionEx(OsVers)
'    winOS = OsVers.dwPlatform
'End Sub
'
