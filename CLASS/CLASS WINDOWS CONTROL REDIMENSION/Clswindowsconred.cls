VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsConRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''******************************************************************************
''***    Delta Copyright                                                             (31/10/2000)  ***
''******************************************************************************
''***    CLASS:                                                                                             ***
''***        clsWindowsConRed                                                                               ***
''******************************************************************************
''***    FONCTION:                                                                                       ***
''***      ----------    Pour La Gestion De  Windows  --------                                  ***
''******************************************************************************
''***    PROGRAMMEUR:                                                                              ***
''***      Royer Jean-Laurent                                                                         ***
''******************************************************************************
'
''******************************************************************************
''***    MODIF :                                                                                            ***
''***      Version 1.0 : 30/10/2000                                                                  ***
''******************************************************************************
'Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer
'
''******************************************************************************
''***    Declaration De Constante Privee                                                       ***
''******************************************************************************
'
'Private Const SWP_DRAWFRAME = &H20
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOSIZE = &H1
'Private Const SWP_NOZORDER = &H4
'Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
'
'Private Const GWL_STYLE = (-16)
'Private Const WS_THICKFRAME = &H40000
'
'
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


