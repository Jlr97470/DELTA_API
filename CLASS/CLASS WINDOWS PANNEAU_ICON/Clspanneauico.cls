VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPanneauIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'******************************************************************************
'***    Delta Copyright                                                             (31/05/2001)  ***
'******************************************************************************
'***    CLASS:                                                                                             ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'**        - Creation initial de la classe                                                          ***
'******************************************************************************
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille en erreur                    ***
'******************************************************************************
Private Const LOGFEUILLENOM = "ClsPanneauIco"                                          ' Je me rapelle le nom de la Feuille

'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini                                                                         ***
'******************************************************************************
Public Enum FENETRECONSTANTES
   GWL_WNDPROC = (-4)
   GWL_USERDATA = (-21)
End Enum

'******************************************************************************
'***    Constante Qui Defini                                                                         ***
'******************************************************************************
Public Enum MESSAGECONSTANTES
   MSG_ADD = &H0
   MSG_MODIFY = &H1
   MSG_DELETE = &H2
End Enum

'******************************************************************************
'***    Constante Qui Defini                                                                         ***
'******************************************************************************
Public Enum FENETREMESSAGECONSTANTES
   WM_USER = &H400
   WM_MOUSEMOVE = &H200
   WM_LBUTTONUP = &H202
   WM_RBUTTONUP = &H205
   WM_MBUTTONUP = &H208
End Enum

'******************************************************************************
'***    Constante Qui Defini                                                                         ***
'******************************************************************************
Public Enum ICONCONSTANTES
   ICON_IDLIST = &H0
   ICON_ASSOCCHANGED = &H8000000
End Enum

'******************************************************************************
'***    Constante Qui Defini                                                                         ***
'******************************************************************************
Public Enum ICONINFORMATIONCONSTANTES
   NIF_MESSAGE = &H1
   NIF_ICON = &H2
   NIF_TIP = &H4
End Enum


'Private Const TRAY_CALLBACK = (WM_USER + 1001&)
'
'Public Sub AddIconToTray(inWinState As Integer, inFormhwnd As Long, inFormIconHandle As Long, inTip As String)
'    theFormWinState = inWinState
'    theWindProc = SetWindowLong(inFormhwnd, GWL_WNDPROC, AddressOf NewWindowProc)
'
'    With typNID
'        .cbSize = Len(typNID)
'        .hwnd = inFormhwnd
'        .uID = 0
'        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
'        .uCallbackMessage = TRAY_CALLBACK
'        .hIcon = inFormIconHandle          '  Or just .hIcon = inForm.Icon
'        .szTip = inTip & vbNullChar
'    End With
'    Shell_NotifyIconA MSG_ADD, typNID
'End Sub
'
'Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, _
'        ByVal wParam As Long, ByVal lParam As Long) As Long
'    If Msg = TRAY_CALLBACK Then
'         If lParam = WM_LBUTTONUP Then           ' Left click icon
'             If theForm.WindowState = vbMinimized Then
'                  theForm.WindowState = vbNormal
'             Else
'                  theForm.WindowState = theFormWinState
'             End If
'             theForm.Show
'             theForm.SetFocus
'             Exit Function
'         End If
'         If lParam = WM_RBUTTONUP Then
'             theForm.PopupMenu theMenu
'             Exit Function
'         End If
'    End If
'    NewWindowProc = CallWindowProc(theWindProc, hwnd, Msg, wParam, lParam)
'End Function
'
'
'
'Public Sub RemoveIconFromTray()
'    With typNID
'        .uFlags = 0
'    End With
'    Shell_NotifyIconA MSG_DELETE, typNID
'    SetWindowLong theForm.hwnd, GWL_WNDPROC, theWindProc
'End Sub
'
'
'
'Public Sub ErrMsgProc(mMsg As String)
'    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
'End Sub
'
'
