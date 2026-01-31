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
'Option Explicit
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
'Private Declare Function GetWindowLong _
'    Lib "user32" _
'    Alias "GetWindowLongA" _
'    (ByVal hwnd As Long, ByVal nIndex As Long) _
'    As Long
'
'Private Declare Function SetWindowLong _
'    Lib "user32" _
'    Alias "SetWindowLongA" _
'    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
'    As Long
'
'
'Private Declare Function SetWindowPos _
'    Lib "user32" _
'    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
'    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
'    ByVal cy As Long, ByVal wFlags As Long) _
'    As Long
'
'
