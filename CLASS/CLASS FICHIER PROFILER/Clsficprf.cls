VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFicPrf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2002
Attribute VB_Description = "PrjDllDeltaApi.clsFicPrf"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'******************************************************************************
'***    Delta Copyright                                                            (31/10/2000)   ***
'******************************************************************************
'***    CLASS:                                                                                             ***
'***        clsFicPrf                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      -------------     Pour La Gestion D'Un Fichier Profiler    -----------------------  ***
'***        1 - Lecture D'Une Donner Dans Un Fichier Profiler                          ***
'***        2 - Ecriture D'Une Donner Dans Un Fichier Profiler                          ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'**        - Creation Initial de la classe                                                          ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille en erreur                    ***
'******************************************************************************
Private Const LOGFEUILLENOM = "clsFicPrf"                                                    ' Je me rapelle le nom de la Feuille


'******************************************************************************
'***    Declaration De Fonction Priver Externe                                             ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                      ***
'***      PrfLectureValeur                                                                            ***
'***      (ByVal strPrfSection As String, ByVal strPrfCle As String,              ***
'***      ByVal strPrfValeurPrftial As String,                                                ***
'***      ByVal strPrfValeur As String,                                                         ***
'***      ByVal lngPrfValeurLongueur As Long,                                              ***
'***      ByVal strPrfFichier As String) As Long                                            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire la valeur d'une cle d'une section du fichier Profiler             ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strPrfSection                         - La Section Du Fichier Prf                  ***
'***      strPrfCle                               - La Cle Du Fichier Prf                        ***
'***      strPrfValeurDefaut                 - La Valeur Prftiale                            ***
'***      strPrfValeur                            - La Valeur Tampon                           ***
'***      lngPrfValeurLongueur              - La Longueur De La Valeur                ***
'***                                                     Tampon                                           ***
'***      strPrfFichier                           - Le Nom Du Fichier Prf                      ***
'***    SORTIE:                                                                                           ***
'***      PrfLectureValeur                    - La Longueur  De La Valeur               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =PrfLectureValeur(strPrfSection, strPrfCle, vbnullstring,                 ***
'***        strPrfValeur, Len(strPrfValeur), strPrfFichier)                                ***
'******************************************************************************
Private Declare Function PrfLectureValeur Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strPrfSection As String, ByVal strPrfCle As String, ByVal strPrfValeurDefaut As String, ByVal strPrfValeur As String, ByVal lngPrfValeurLongueur As Long, ByVal strPrfFichier As String) As Long
                                                                                                                                                                                        
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      PrfEcritureValeur                                                                            ***
'***     (ByVal strPrfSection As String, ByVal strPrfCle As String,                ***
'***     ByVal anyPrfValeur As Any, ByVal strPrfFichier As String)                ***
'****    As Boolean                                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire la valeur d'une cle d'une section du fichier Profiler         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strPrfSection                          - La Section Du Fichier Prf                 ***
'***      strPrfCle                               - La Cle Du Fichier Prf                         ***
'***      anyPrfValeur                           - La Valeur De La Cle                        ***
'***      strPrfFichier                          - Le Nom Du Fichier Prf                      ***
'***    SORTIE:                                                                                           ***
'***      PrfEcritureValeur                 - La Longueur  De La Valeur                  ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =PrfEcritureValeur(ByVal strPrfSection As String,                            ***
'***        ByVal strPrfCle As string,                                                             ***
'***        ByVal anyPrfValeur As Any, ByVal strPrfFichier As String)            ***
'******************************************************************************
Private Declare Function PrfEcritureValeur Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strPrfSection As String, ByVal strPrfCle As String, ByVal anyPrfValeur As Any, ByVal strPrfFichier As String) As Boolean

'******************************************************************************
'***    Declaration De Fonction Public                                                         ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      LectureValeur(ByVal strPrfSection As String,                                   ***
'***      ByVal strPrfCle As String,                                                               ***
'***      ByVal strPrfValeurDefaut As String, ByVal strPrfFichier As String)   ***
'***      As String                                                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire une valeur de cle dans une section du fichier Prf                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strPrfFichier                              - Le nom du fichier                         ***
'***      strPrfSection                             - Le nom de la section                   ***
'***      strPrfCle                                    - La cle de la section                     ***
'***      strPrfValeur                               - La valeur De La Cle                     ***
'***      strPrfValeurDefaut                     - La valeur  initiale De La Cle         ***
'***    SORTIE:                                                                                           ***
'***      strPrfValeur                               - La valeur De La Cle                     ***
'***      EcritureValeur                           - Renvoie la validiter de l'operation***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =LectureValeur("FICHIER.INI","COULEUR","NOIR","10",strValeur)   ***
'******************************************************************************
' ##SUMMARY Lire une valeur dans le fichier profiler
' ##RETURNS Si la cle a ete trouver
Public Function LectureValeur(ByVal strPrfFichier As String, ByVal strPrfSection As String, ByVal strPrfCle As String, ByRef strPrfValeur As String, ByVal strPrfValeurDefaut As String) As Boolean
' ##PARAM strPrfFichier Nom du fichier
' ##PARAM strPrfSection Nom de la section
' ##PARAM strPrfCle Nom de la cle
' ##PARAM strPrfValeur Valeur de la cle
' ##PARAM strPrfValeurDefaut Valeur par defaut de la cle
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureValeur_Erreur
   ' J'Prftialise la variable tampon
   strPrfValeur = String$(2048, " ")
   'Je recupere la valeur
   strPrfValeur = Left$(strPrfValeur, PrfLectureValeur(strPrfSection, ByVal strPrfCle, strPrfValeurDefaut, strPrfValeur, Len(strPrfValeur), strPrfFichier))
   ' Je regarde la valeur
   LectureValeur = (strPrfValeur <> vbNullString)
   ' Fin
LectureValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
LectureValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, FEUILLECLASS, LOGFEUILLENOM, LIBELLEFONCTION, "LectureValeur", _
      strPrfSection & "," & strPrfCle & "," & strPrfValeurDefaut & "," & strPrfFichier, Err
   'Je renvoie que la valeur n'a pas ete lu
   LectureValeur = False
   ' Je Continue
   Resume LectureValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      EcritureValeur(ByVal strPrfFichier As String,                                   ***
'***      ByVal strPrfSection As String,                                                         ***
'***      ByVal  strPrfCle As String,ByVal strPrfValeur As String) As Boolean***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire une valeur de cle dans une section                                 ***
'***      du fichier Prf                                                                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strPrfFichier                         - Le nom du fichier                              ***
'***      strPrfSection                         - Le nom de la section                        ***
'***      strPrfCle                               - La cle de la section                          ***
'***      strPrfValeur                          - La valeur de la cle                            ***
'***    SORTIE:                                                                                           ***
'***      EcritureValeur                       - Renvoie la validiter de l'operation    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EcritureValeur("COULEUR","NOIR","10","FICHIER.INI")                  ***
'******************************************************************************
Public Function EcritureValeur(ByVal strPrfFichier As String, ByVal strPrfSection As String, ByVal strPrfCle As String, ByVal strPrfValeur As String) As Boolean
' ##PARAM strPrfFichier Nom du fichier
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureValeur_Erreur
   ' J'ecrit la valeur dans le fichier Prf
   EcritureValeur = PrfEcritureValeur(strPrfSection, strPrfCle, strPrfValeur, strPrfFichier)
   ' Fin
EcritureValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EcritureValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, FEUILLECLASS, LOGFEUILLENOM, LIBELLEFONCTION, "EcritureValeur", _
      strPrfSection & "," & strPrfCle & "," & strPrfValeur & "," & strPrfFichier, Err
   'Je renvoie que la valeur n'a pas ete ecrit
   EcritureValeur = False
   ' Je Continue
   Resume EcritureValeur_Exit
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
   gfloLogDeltaApi.AjouteErreur App, FEUILLECLASS, LOGFEUILLENOM, LIBELLEPROCEDURE, "Class_Initialize", vbNullString, Err
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
   gfloLogDeltaApi.AjouteErreur App, FEUILLECLASS, LOGFEUILLENOM, LIBELLEPROCEDURE, "Class_Terminate", vbNullString, Err
   ' Je Continue
   Resume Class_Terminate_Exit
   ' Fin
End Sub

