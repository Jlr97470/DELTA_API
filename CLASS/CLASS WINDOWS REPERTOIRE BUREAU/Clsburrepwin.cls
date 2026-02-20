VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepBurWin"
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
'***        clsRepBubWin                                                                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      ----------    Pour La Gestion De  Windows  --------                                  ***
'***        - Pour La Gestion Des Repertoires Speciaux De Windows               ***
'***        - Recuperer les repertoires speciaux de windows                           ***
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

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille en erreur                    ***
'******************************************************************************
Private Const mconFEUILLETYPE = "CLASS"                                                      ' La feuille du source
Private Const mconFEUILLENOM = "clsBurRepWin"                                            ' Je me rapelle le nom de la Feuille

'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Les numeros de repertoires specials                                 ***
'******************************************************************************
Public Enum BUREAUREPERTOIREIDCONSTANTES
    CSIDL_DESKTOP = &H0&                                                                           ' ..\WinNT\profiles\username\Desktop
    CSIDL_PROGRAMS = &H2&                                                                        ' ..\WinNT\profiles\username\Start Menu\Programs
    CSIDL_CONTROLS = &H3&                                                                         ' No Path
    CSIDL_PRINTERS = &H4&                                                                          ' No Path
    CSIDL_PERSONAL = &H5&                                                                         ' ..\WinNT\profiles\username\Personal
    CSIDL_FAVORITES = &H6&                                                                        ' ..\WinNT\profiles\username\Favorites
    CSIDL_STARTUP = &H7&                                                                           ' ..\WinNT\profiles\username\Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&                                                                             ' ..\WinNT\profiles\username\Recent
    CSIDL_SENDTO = &H9&                                                                             ' ..\WinNT\profiles\username\SendTo
    CSIDL_BITBUCKET = &HA&                                                                       ' No Path
    CSIDL_STARTMENU = &HB&                                                                      ' ..\WinNT\profiles\username\Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10&                                                       ' ..\WinNT\profiles\username\Desktop
    CSIDL_DRIVES = &H11&                                                                            ' No Path
    CSIDL_NETWORK = &H12&                                                                        ' No Path
    CSIDL_NETHOOD = &H13&                                                                        ' ..\WinNT\profiles\username\NetHood
    CSIDL_FONTS = &H14&                                                                             ' ..\WinNT\fonts
    CSIDL_TEMPLATES = &H15&                                                                     ' ..\WinNT\ShellNew
    CSIDL_COMMON_STARTMENU = &H16&                                                    ' ..\WinNT\profiles\All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&                                                      ' ..\WinNT\profiles\All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&                                                         ' ..\WinNT\profiles\All Users\Start Menu\Programs\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19&                                       '..\WinNT\profiles\All Users\Desktop
    CSIDL_APPDATA = &H1A&                                                                         ' ..\WinNT\profiles\username\Application Data
    CSIDL_PRINTHOOD = &H1B&                                                                     ' ..\WinNT\profiles\username\PrintHood
End Enum

'******************************************************************************
'***    Declaration De Fonction Private Externe                                            ***
'******************************************************************************
    
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function BurLectureRepertoireID Lib "Shell32" Alias "SHGetSpecialFolderLocation" (ByRef anyNull As Any, ByVal intRepertoireID As Integer, lngRepertoireIDEntete As Long) As Long
    
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function BurLectureRepertoire Lib "Shell32" Alias "SHGetPathFromIDListW" (ByVal lngRepertoireIDEntete As Long, ByVal strRepertoire As Long) As Long

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function LectureRepertoire(ByVal intRepertoireID As BUREAUREPERTOIREIDCONSTANTES, ByRef strRepertoire As String) As Boolean
    Dim lngBurOperation As Long
    Dim lngRepertoireIDEntete As Long
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo LectureRepertoire_Erreur
       
    lngBurOperation = BurLectureRepertoireID(Null, intRepertoireID, lngRepertoireIDEntete)
    
    If (lngBurOperation = 0) Then
    
        strRepertoire = Space(1024)
    
        lngBurOperation = BurLectureRepertoire(lngRepertoireIDEntete, StrPtr(strRepertoire))
        
        If (lngBurOperation = 1) Then
        
            strRepertoire = Left(strRepertoire, InStr(1, strRepertoire, Chr$(0)) - 1)
            
            LectureRepertoire = True
            
        End If
    End If
   ' Fin
LectureRepertoire_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
LectureRepertoire_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROCEDURE, "LectureRepertoire", CStr(intRepertoireID) & "," & strRepertoire, Err
   ' Je Continue
   Resume LectureRepertoire_Exit
   ' Fin
End Function

'******************************************************************************
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***     Class_Initialize                                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour creer une collection de cle a la creation                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     Neant                                                                                              ***
'***    SORTIE:                                                                                           ***
'***     Neant                                                                                              ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     Dim objObject As New Class                                                           ***
'******************************************************************************
Private Sub Class_Initialize()
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Class_Initialize_Erreur
   ' Fin
Class_Initialize_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
Class_Initialize_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROCEDURE, "Class_Initialize", vbNullString, Err
   ' Je Continue
   Resume Class_Initialize_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***     Class_Terminate                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer la collection de cle a la suppression                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     Neant                                                                                              ***
'***    SORTIE:                                                                                           ***
'***     Neant                                                                                              ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     Class=Nothing                                                                                 ***
'******************************************************************************
Private Sub Class_Terminate()
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Class_Terminate_Erreur
   ' Fin
Class_Terminate_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
Class_Terminate_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROCEDURE, "Class_Terminate", vbNullString, Err
   ' Je Continue
   Resume Class_Terminate_Exit
   ' Fin
End Sub


