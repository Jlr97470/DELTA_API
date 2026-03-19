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
'******************************************************************************
'***    Delta Copyright                                                             (31/10/2000)  ***
'******************************************************************************
'***    CLASS:                                                                                             ***
'***        clsWindowsConRed                                                                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      ----------    Pour La Gestion De  Windows  --------                                  ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000                                                                  ***
'******************************************************************************
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************


Private Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

Private OsVers As OsVersionInfo

Private Const Win95 = 1&

Private Const WinNT = 2&

Private winOS&

Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpStruct As OsVersionInfo)

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long


Public Function GetVersion32(Index As Integer) As Variant
    Dim tempLong

    OsVers.dwVersionInfoSize = 148&
    
    tempLong = GetVersionEx(OsVers)
    
    Select Case Index
        Case 0
        
            GetVersion32 = OsVers.dwMajorVersion
            
        Case 0
        
            GetVersion32 = OsVers.dwMinorVersion
        
        Case 0
        
            GetVersion32 = OsVers.dwBuildNumber
    
        Case 0
        
            GetVersion32 = OsVers.dwPlatform
    
        Case 0
        
            GetVersion32 = OsVers.szCSDVersion
    
    End Select
    
End Function

