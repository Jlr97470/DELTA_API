VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsFen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
'Option Explicit
'Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
'Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
'Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'
'Private Const GW_HWNDFIRST = 0
'Private Const GW_HWNDLAST = 1
'Private Const GW_HWNDNEXT = 2
'Private Const GW_HWNDPREV = 3
'Private Const GW_OWNER = 4
'Private Const GW_CHILD = 5
'Private Const GW_MAX = 5
'
'Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Function FindWindowPartial(ByVal TitlePart As String) As Long
''Borrowed from VBPJ and ported to 32-bit
'
'     Dim hWndTmp As Long
'     Dim nRet As Integer
'     Dim TitleTmp As String
'
'     'We alter the title to compare it case-insensitively.
'     TitlePart = UCase$(TitlePart)
'
'     'First find all the open windows so we can
'     'loop through them and find the right one.
'     hWndTmp = FindWindow(0&, 0&)
'
'     Do Until hWndTmp = 0
'          TitleTmp = Space$(256)
'          nRet = GetWindowText(hWndTmp, TitleTmp, Len(TitleTmp))
'          If nRet Then
'               'Let's prepare to compare
'               TitleTmp = UCase$(VBA.Left$(TitleTmp, nRet))
'               'Now we see if the window we chased down actually
'               'has the caption we want.
'               If InStr(TitleTmp, TitlePart) Then
'                    FindWindowPartial = hWndTmp
'                    Exit Do
'               End If
'          End If
'          hWndTmp = GetWindow(hWndTmp, GW_HWNDNEXT)
'     Loop
'End Function
'
'
