VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsSou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'DefInt A-Z
'
''  Mouse/cursor functions.
'
'Private lShowCursor As Long
'
''Present
'Private Const SM_MOUSEPRESENT = 19
'Private Const SM_SWAPBUTTON = 23
'Private Const SM_CMOUSEBUTTONS = 43
'
''mouse position
'Private Type POINTAPI_TYPE
'    x As Long
'    Y As Long
'End Type
'
''clip
'Private Type RECT_TYPE
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI_TYPE) As Long
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'
''visible
'Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'
''buttons down
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT_TYPE) As Long
'Private Declare Function ClipCursor Lib "user32" (lpRect As RECT_TYPE) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'
'' swap
'Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'
''
''  Hides the mouse cursor.
''
'Public Sub HideMouse()
'   Dim result As Integer
'
'   Do
'      lShowCursor = lShowCursor - 1
'      result = ShowCursor(False)
'   Loop Until result < 0
'
'End Sub
'
''
''  Restores the mouse cursor to it's previous state regardless
''  if HideMouse and ShowMouse were called.
''
'Public Sub RestoreMouse()
'   If lShowCursor > 0 Then
'      Do While lShowCursor <> 0
'         ShowCursor (False)
'         lShowCursor = lShowCursor - 1
'      Loop
'   ElseIf lShowCursor < 0 Then
'      Do While lShowCursor <> 0
'         ShowCursor (True)
'         lShowCursor = lShowCursor + 1
'      Loop
'   End If
'End Sub
'
'
''
''  Show's the mouse cursor.
''
'Public Sub ShowMouse()
'   Dim result
'
'   Do
'      lShowCursor = lShowCursor - 1
'      result = ShowCursor(True)
'   Loop Until result >= 0
'
'End Sub
'
'Public Sub ButtonsSwap()
'
'    SwapMouseButton True
'
'End Sub
'
'Public Property Get CurrentX() As Long
'
'    Dim pt As POINTAPI_TYPE
'    GetCursorPos pt
'    CurrentX = pt.x
'
'End Property
'Public Property Let CurrentX(x As Long)
'
'    Dim pt As POINTAPI_TYPE
'
'    GetCursorPos pt
'    pt.x = x
'    SetCursorPos pt.x, pt.Y
'
'End Property
'
'Public Property Let CurrentY(Y As Long)
'
'    Dim pt As POINTAPI_TYPE
'
'    GetCursorPos pt
'    pt.Y = Y
'    SetCursorPos pt.x, pt.Y
'
'End Property
'
'
'Public Property Get CurrentY() As Long
'
'    Dim pt As POINTAPI_TYPE
'    GetCursorPos pt
'    CurrentY = pt.Y
'
'End Property
'
'
'Public Property Get Visible() As Boolean
'
'    Dim oldmode As Integer
'
'    oldmode = ShowCursor(True)
'    ShowCursor False
'    Visible = oldmode
'
'End Property
'
'Public Property Let Visible(ByVal bNewValue As Boolean)
'
'    Dim oldmode As Integer
'
'    If bNewValue Then
'        oldmode = -1
'        While (oldmode < 1)
'            oldmode = ShowCursor(True)
'        Wend
'    Else
'        oldmode = 1
'        While (oldmode >= 0)
'            oldmode = ShowCursor(False)
'        Wend
'    End If
'
'End Property
'
'Public Function IsPresent() As Boolean
'
'    IsPresent = GetSystemMetrics(SM_MOUSEPRESENT) = 1
'
'End Function
'
'Public Sub ButtonsUnswap()
'
'    SwapMouseButton False
'
'End Sub
'
'Public Function ButtonsSwapped() As Boolean
'
'    ButtonsSwapped = GetSystemMetrics(SM_SWAPBUTTON) = 1
'
'End Function
'
'Public Function ButtonDownLeft() As Boolean
'
'    GetAsyncKeyState vbKeyLButton
'    ButtonDownLeft = Not (GetAsyncKeyState(vbKeyLButton) And &HFFFF) = 0
'
'End Function
'Public Function ButtonDownMiddle() As Boolean
'
'    GetAsyncKeyState vbKeyMButton
'    ButtonDownMiddle = Not (GetAsyncKeyState(vbKeyMButton) And &HFFFF) = 0
'
'End Function
'
'Public Function ButtonDownRight() As Boolean
'
'    GetAsyncKeyState vbKeyRButton
'    ButtonDownRight = Not (GetAsyncKeyState(vbKeyRButton) And &HFFFF) = 0
'
'End Function
'
'
'Public Sub Move(x As Long, Y As Long)
'
'    SetCursorPos x, Y
'
'End Sub
'
'Public Sub ClipToControl(ByVal ctl As Object)
'
'    Dim hwnd As Long
'    Dim typ_RECT As RECT_TYPE
'
'    hwnd = ctl.hwnd
'
'    GetWindowRect hwnd, typ_RECT
'    SetCursorPos typ_RECT.Left + ((typ_RECT.Right - typ_RECT.Left) / 2), _
'                 typ_RECT.Top + ((typ_RECT.Bottom - typ_RECT.Top) / 2)
'
'    ClipCursor typ_RECT
'
'myExit:
'
'End Sub
'
'Public Sub ClipToDesktop()
'
'    Dim typ_RECT As RECT_TYPE
'
'    GetWindowRect GetDesktopWindow(), typ_RECT
'    ClipCursor typ_RECT
'
'End Sub
'Public Sub ClipToRectangle(Left As Long, Top As Long, Width As Long, Height As Long)
'
'    Dim typ_RECT As RECT_TYPE
'
'    typ_RECT.Left = Left
'    typ_RECT.Top = Top
'    typ_RECT.Right = Left + Width
'    typ_RECT.Bottom = Top + Height
'
'    SetCursorPos typ_RECT.Left + ((typ_RECT.Right - typ_RECT.Left) / 2), _
'                 typ_RECT.Top + ((typ_RECT.Bottom - typ_RECT.Top) / 2)
'
'    ClipCursor typ_RECT
'
'End Sub
'
'Private Sub Class_Terminate()
'
'    'Clean up any possible bad things we did here
'    ClipToDesktop   'remove clip
'    Visible = True  'make it visible
'
'End Sub
'
'Public Function ButtonCount() As Long
'
'    ButtonCount = GetSystemMetrics(SM_CMOUSEBUTTONS)
'
'End Function
'
'Public Function IsOver() As Long
'
'    Dim lppoint As POINTAPI_TYPE
'    GetCursorPos lppoint
'    IsOver = WindowFromPoint(lppoint.Y, lppoint.x)
'
'End Function
'
'
'
