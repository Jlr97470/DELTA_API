VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFicLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2005
Attribute VB_Description = "PrjDllDeltaApi.clsFicLog"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'******************************************************************************
'***    Delta Copyright                                                             (03/06/2001)  ***
'******************************************************************************
'***    CLASS:                                                                                             ***
'***        clsFicLog                                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      ------------       La Gestion D'Un Fichier journal formater      --------------   ***
'***        1 - Lire Un Enregistrement Dans Le Fichier Journal                        ***
'***        2 - Ecrire Un Enregistrement Dans Le Fichier Journal                     ***
'***        3 - Deplacement Dans Les Enregistremnt Du Fichier Journal          ***
'***        4 - Ajouter Les Informations D'Une Erreur                                      ***
'***        5 - Afficher Les Informations D'Une Erreur                                     ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'**        - Creation initial de la classe                                                          ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Le nom de la Feuille                                          ***
'******************************************************************************
Private Const LOGFEUILENOM = "clsFicLog"

'******************************************************************************
'***    Constante Qui Defini Les Libelles D'Un Enregistrement                     ***
'******************************************************************************
Private Const LOGAPPLICATION = "APPLICATION"
Private Const LOGVERSION = "VERSION"
Private Const LOGFEUILLE = "FEUILLE"
Private Const LOGTYPE = "TYPE"
Private Const LOGNOM = "NOM"
Private Const LOGINSTRUCTION = "INSTRUCTION"
Private Const LOGPROPRIETE = "PROPRIETE"
Private Const LOGFONCTION = "FONCTION"
Private Const LOGPROCEDURE = "PROCEDURE"
Private Const LOGEVENEMENT = "EVENEMENT"
Private Const LOGPARAMETRE = "PARAMETRE"
Private Const LOGERREUR = "ERREUR"
Private Const LOGNUMERO = "NUMERO"
Private Const LOGDESCRIPTION = "DESCRIPTION"
Private Const LOGJOUR = "JOUR"
Private Const LOGHEURE = "HEURE"

'******************************************************************************
'***    Declaration Constante Public                                                            ***
'******************************************************************************

'******************************************************************************
'***    Constante De Type De Feuille En Erreur                                            ***
'******************************************************************************
Public Enum LOGFEUILLETYPETYPEENUM
   FEUILLEFORM
   FEUILLEMODULE
   FEUILLECLASS
   FEUILLEMDIFORM
End Enum

'******************************************************************************
'***    Constante De Type D'Instruction En Erreur                                        ***
'******************************************************************************
Public Enum LOGINSTRUCTIONTYPEENUM
   INSTRUCTIONFONCTION
   INSTRUCTIONPROCEDURE
   INSTRUCTIONEVENEMENT
   INSTRUCTIONPROPRIETER
End Enum

'******************************************************************************
'***    Declaration De Type Priver                                                               ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion D'Un Fichier En Acces Sequentiels Indexer       ***
'******************************************************************************
Private Type LOGPOSITIONTYPE
   intPositionIndex As Integer                                                                         ' Le numero de l'enregistrement en cours
   intPositionMax As Integer                                                                           ' Le nombre d'enregistremnt
   intPositionFichierEntete As Integer                                                             ' L'entete de fichier ouvert
End Type

'******************************************************************************
'***    Type Pour La Gestion D'Un Enregistrement Du Fichier Log                 ***
'******************************************************************************
Private Type LOGENRTYPE
   strEnrLigne(21) As String * 256                                                                     ' Les lignes de l'enregistrement
End Type

'******************************************************************************
'***    Type Pour La Gestion De La Version De La Dll Delta Api                    ***
'******************************************************************************
Private Type LOGDELTAAPITYPE
   strDeltaApiVersion As String                                                                      ' La version de la dll deltaapi
End Type

'******************************************************************************
'***    Type Pour La Gestion De L'Application                                              ***
'******************************************************************************
Private Type LOGAPPLICATIONTYPE
   strApplicationNom As String                                                                      ' Le nom de l'application
   strApplicationVersion As String                                                                  ' La version de l'application
End Type

'******************************************************************************
'***    Type Pour La Gestion De La Feuille                                                   ***
'******************************************************************************
Private Type LOGFEUILLETYPE
   intFeuilleType As String                                                                            ' Le type de la feuille en erreur
   strFeuilleNom As String                                                                             ' Le nom de la feuille en erreur
End Type

'******************************************************************************
'***    Type Pour La Gestion De L'Instruction                                               ***
'******************************************************************************
Private Type LOGINSTRUCTIONTYPE
   IntInstructionType As Integer                                                                     ' Le type de la Instruction en erreur
   strInstructionNom As String                                                                       ' Le nom de la Instruction en erreur
   strInstructionParametre As String                                                              ' Les parametres de la Instruction
End Type

'******************************************************************************
'***    Type Pour La Gestion De L'Erreur                                                      ***
'******************************************************************************
Private Type LOGERREURTYPE
   strErreurNumero As String                                                                         ' Le numero de l'erreur
   strErreurDescription As String                                                                   ' La description de l'erreur
End Type

'******************************************************************************
'***    Type Pour La Gestion Du Jour Et L'Heure                                           ***
'******************************************************************************
Private Type LOGDATETYPE
   strDateJour As String                                                                               ' Le jour
   strDateHeure As String                                                                             ' L'heure
End Type

'******************************************************************************
'***    Type Pour La Gestion Du Fichier Log                                                 ***
'******************************************************************************
Private Type LOGFICHIERTYPE
   ulinFichierIndex As LOGPOSITIONTYPE                                                             ' La position dans le fichier
   ulenFichierEnr As LOGENRTYPE                                                                        ' Un enregistrement du fichier
   uldaFichierDeltaApi As LOGDELTAAPITYPE                                                       ' La version de la dll de l'enregistrement
   ulapFichierApplication As LOGAPPLICATIONTYPE                                              ' L'application de l'enregistrement
   ulfeFichierFeuille As LOGFEUILLETYPE                                                              ' La feuille de l'enregistrement
   ulinFichierInstruction As LOGINSTRUCTIONTYPE                                               ' L'instruction de l'enregistrement
   ulerFichierErreur As LOGERREURTYPE                                                              ' L'erreur de l'enregistrement
   uldtFichierDate As LOGDATETYPE                                                                    ' La date de l'enregistrement
End Type

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour La Gestion Du Fichier Journal                                       ***
'******************************************************************************
Private mulfiLogFichier As LOGFICHIERTYPE                                                          ' Un object fichier journal formater

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Lecture() As Boolean                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lire un enregistrement du fichier journal                                  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Lecture - Renvoie si l'enregistrement a ete lu                                  ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Lecture()                                                                                       ***
'******************************************************************************
Public Function Lecture() As Boolean
Attribute Lecture.VB_Description = "******************************************************************************"
Attribute Lecture.VB_HelpID = 2006
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Lecture_Erreur
   ' Je vais
   With mulfiLogFichier
      ' Je regarde le numero d'enregistrement
      Select Case .ulinFichierIndex.intPositionIndex
          Case vbEmpty
               ' Le numero est a zero
               ' Je renvoie que je n'est pas lu l'enregistrement
               Lecture = False
          Case Else
               ' Le numero n'est pas a zero
               ' Je lit l'enregistrement
               Get .ulinFichierIndex.intPositionFichierEntete, .ulinFichierIndex.intPositionIndex, .ulenFichierEnr
               ' Je lit les donnes de la dll
               LectureDeltaApiVersion
               ' Je lit les donnes de l'application
               LectureApplication
               ' Je lit les donnees de la feuille
               LectureFeuille
               ' Je lit les donnees de la Instruction
               LectureInstruction
               ' Je lit les donnees de l'erreur
               LectureErreur
               ' Je lit les donnees de la date
               LectureDate
               ' Je renvoie que j'ai bien lu l'enregistrement
               Lecture = True
      End Select
   End With
   ' Fin
Lecture_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Lecture_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Lecture", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je n'est pas lu l'enregistrement
   Lecture = False
   ' Je Continue
   Resume Lecture_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Ecriture() As Boolean                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire l'enregistrement du fichier log                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Ecriture - Renvoie si l'enregistrement a ete ecrit                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Ecriture()                                                                                      ***
'******************************************************************************
Public Function Ecriture() As Boolean
Attribute Ecriture.VB_Description = "En Cas D'Erreur Je Gere L'Erreur"
Attribute Ecriture.VB_HelpID = 2007
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Ecriture_Erreur
   ' Je renvoie que j'ai bien ecrit l'enregistrement
   Ecriture = True
   ' Je vais
   With mulfiLogFichier
      ' Je regarde le numero d'enregistrement
      Select Case .ulinFichierIndex.intPositionIndex
         Case vbEmpty
              ' Le numero est a zero
              ' Je renvoie que je n'ai pas ecrit l'enregistrement
              Ecriture = False
         Case Else
              ' Le numero n'est pas a zero
              ' Je vais
              With .ulenFichierEnr
                    ' Je met une chaine de separation
                    .strEnrLigne(0) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(2) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(4) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(5) = "***" & LOGFEUILLE & ":" & Space(247 - Len(LOGFEUILLE)) & "***" & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(7) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(8) = "***" & LOGINSTRUCTION & ":" & Space(247 - Len(LOGINSTRUCTION)) & "***" & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(10) = "***" & LOGPARAMETRE & ":" & Space(247 - Len(LOGPARAMETRE)) & "***" & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(13) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(15) = "***" & LOGDESCRIPTION & ":" & Space(247 - Len(LOGDESCRIPTION)) & "***" & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(19) = String(254, "*") & vbCrLf
                    ' Je met une chaine de separation
                    .strEnrLigne(21) = String(254, "*") & vbCrLf
              End With
              ' J'ecrit les donnes de la version de la dll
              EcritureDeltaApiVersion
              ' J'ecrit les donnes de l'application
              EcritureApplication
              ' J'ecrit les donnees de la feuille
              EcritureFeuille
              ' J'ecrit les donnes de la Instruction
              EcritureInstruction
              ' J'ecrit les donnes du de l'erreur
              EcritureErreur
              ' J'ecrit les donnes de la date de l'erreur
              EcritureDate
              ' J'ecrit l'enregistrement
              Put .ulinFichierIndex.intPositionFichierEntete, .ulinFichierIndex.intPositionIndex, .ulenFichierEnr
      End Select
   End With
   ' Fin
Ecriture_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Ecriture_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Ecriture", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je n 'ai pas ecrit l'enregistrement
   Ecriture = False
   ' Je Continue
   Resume Ecriture_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Ouvre(ByVal strFichier As String) As Boolean                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ouvrir un fichier log                                                                 ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strFichier - Le Fichier log                                                                ***
'***    SORTIE:                                                                                           ***
'***      Ouvre        - Renvoie Si Le Fichier A Ete Ouvert                               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Ouvre("ESSAI")                                                                             ***
'******************************************************************************
Public Function Ouvre(ByVal strFichier As String) As Boolean
Attribute Ouvre.VB_Description = "En Cas D'Erreur Je Gere L'Erreur"
Attribute Ouvre.VB_HelpID = 2008
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Ouvre_Erreur
   ' Je ferme le fichier
   Ouvre = Ferme
   ' Je regarde si j'ai bien femer le fichier
   Select Case Ouvre
        Case True
             ' J'ai bien fermer le fichier
             ' Je vais beaucoup utiliser
             With mulfiLogFichier
                With .ulinFichierIndex
                    ' Je defini un numero de fichier
                    .intPositionFichierEntete = FreeFile
                    ' J'ouvre Le Fichier
                    Open strFichier For Random As .intPositionFichierEntete Len = Len(mulfiLogFichier.ulenFichierEnr)
                    ' Je me rapelle le nombre d'enregistrement
                    .intPositionMax = LOF(.intPositionFichierEntete) \ Len(mulfiLogFichier.ulenFichierEnr)
                End With
                ' Je regarde Le nombre d'enregistrement
                Select Case .ulinFichierIndex.intPositionMax
                     Case vbEmpty
                          ' Le nombre d'enregistrement est a zero
                          ' Je met a zero l'index de l'enregistrement
                          .ulinFichierIndex.intPositionIndex = vbEmpty
                          ' J'initialise les variables
                          Initialise
                     Case Else
                          ' Le nombre d'enregistrement n'est pas a zero
                          ' Je me place sur le premier
                          .ulinFichierIndex.intPositionIndex = 1
                          ' Je lit le premier enregistrement
                          Ouvre = Lecture
                End Select
            End With
        Case False
             ' Je n'est pas fermer le fichier
   End Select
   ' Fin
Ouvre_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Ouvre_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Ouvre", strFichier, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je me rapelle que je n'ai pas ouvert le fichier
   Ouvre = False
   ' Je Continue
   Resume Ouvre_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Ferme() As Boolean                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour fermer le fichier log                                                                 ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Ferme - Renvoie Si Le Fichier A Ete Fermer                                     ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Ferme()                                                                                         ***
'******************************************************************************
Public Function Ferme() As Boolean
Attribute Ferme.VB_HelpID = 2009
   Dim lngEnrLigneNum As Long
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Ferme_Erreur
   ' J'ai bien fermer le fichier
   Ferme = True
    ' Je vais beaucoup
    With mulfiLogFichier
        ' Je regarde le numero de fichier log de l'application
        Select Case .ulinFichierIndex.intPositionFichierEntete
             Case vbEmpty
                  ' Le numero est a zero
             Case Else
                  ' Il y a un numero
                  With .ulinFichierIndex
                        ' Je ferme le fichier
                        Close #.intPositionFichierEntete
                        ' Je met a zero le numero de fichier
                        .intPositionFichierEntete = vbEmpty
                        ' Je met a zero le nombre d'enregistrement
                        .intPositionMax = vbEmpty
                        ' Je met a zero l'index de l'enregistrement
                        .intPositionMax = vbEmpty
                    End With
                    ' J'initialise les variables
                    Initialise
        End Select
    End With
   ' Fin
Ferme_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Ferme_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Ferme", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je n'ai pas fermer le fichier
   Ferme = False
   ' Je Continue
   Resume Ferme_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Premier() As Boolean                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Se positionner sur le premier enregistrement                           ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Premier - Renvoie si on est sur le                                                    ***
'***                    premier enregistrement                                                     ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Premier()                                                                                       ***
'******************************************************************************
Public Function Premier() As Boolean
Attribute Premier.VB_HelpID = 2010
   Dim intPositionIndex As Integer                                                                   'Le numero d'enregistrement courant
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Premier_Erreur
   ' Je vais
   With mulfiLogFichier.ulinFichierIndex
        ' Je regarde le nombre d'enregistrement
        Select Case .intPositionMax
             Case vbEmpty
                  ' Le nombre d'enregistrement est a zero
                  ' Je renvoie que je ne suis pas sur le Premier enregistrement
                  Premier = False
             Case Else
                  ' Le nombre d'enregistremnt n'est pas a zero
                  ' Je me souvient de l'ancien numero d'enregistremnt
                  intPositionIndex = .intPositionIndex
                  ' Je n'ai pas lu l'enregistrement
                  .intPositionIndex = 1
                  ' Je lit le Premier enregistrement
                  Premier = Lecture()
                  ' Je regarde si j'ai bien lu le dernier enregistrement
                  Select Case Premier
                        Case True
                             ' J'ai bien lut l'enregistrement
                        Case False
                             ' Je n'ai pas lu l'enregistrement
                             .intPositionIndex = intPositionIndex
                             ' Je lit l'ancien enregistrement
                             Lecture
                   End Select
        End Select
    End With
   ' Fin
Premier_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Premier_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Premier", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je ne suis pas sur le dernier enregistrement
   Premier = False
   ' Je Continue
   Resume Premier_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Dernier() As Boolean                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Se positionner sur le dernier enregistrement                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Dernier - Renvoie si on est sur le                                                    ***
'***                    dernier enregistrement                                                     ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Dernier()                                                                                       ***
'******************************************************************************
Public Function Dernier() As Boolean
Attribute Dernier.VB_HelpID = 2011
   Dim intPositionIndex As Integer                                                                   'Le numero d'enregistrement courant
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Dernier_Erreur
   ' Je vais
   With mulfiLogFichier.ulinFichierIndex
        ' Je regarde le nombre d'enregistrement
        Select Case .intPositionMax
             Case vbEmpty
                  ' Le nombre d'enregistrement est a zero
                  ' Je renvoie que je ne suis pas sur le dernier enregistrement
                  Dernier = False
             Case Else
                  ' Le nombre d'enregistremnt n'est pas a zero
                  ' Je me souvient de l'ancien numero d'enregistremnt
                  intPositionIndex = .intPositionIndex
                  ' Je n'ai pas lu l'enregistrement
                  .intPositionIndex = .intPositionMax
                  ' Je lit le dernier enregistrement
                  Dernier = Lecture()
                  ' Je regarde si j'ai bien lu le dernier enregistrement
                  Select Case Dernier
                        Case True
                             ' J'ai bien lut l'enregistrement
                        Case False
                             ' Je n'ai pas lu l'enregistrement
                             .intPositionIndex = intPositionIndex
                             ' Je lit l'ancien enregistrement
                             Lecture
                   End Select
        End Select
    End With
   ' Fin
Dernier_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Dernier_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Dernier", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je ne suis pas sur le dernier enregistrement
   Dernier = False
   ' Je Continue
   Resume Dernier_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Precedent() As Boolean                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Se positionner sur le Precedent enregistrement                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Precedent - Renvoie si on est sur le                                                ***
'***                    Precedent enregistrement                                                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Precedent()                                                                                   ***
'******************************************************************************
Public Function Precedent() As Boolean
Attribute Precedent.VB_HelpID = 2012
   Dim intPositionIndex As Integer                                                                   'Le numero d'enregistrement courant
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Precedent_Erreur
   ' Je vais
   With mulfiLogFichier.ulinFichierIndex
        ' Je regarde le nombre d'enregistrement
        Select Case .intPositionMax
             Case vbEmpty
                  ' Le nombre d'enregistrement est a zero
                  ' Je renvoie que je ne suis pas sur le Premier enregistrement
                  Precedent = False
             Case Else
                ' Le nombre d'enregistremnt n'est pas a zero
                ' Je me souvient de l'ancien numero d'enregistremnt
                intPositionIndex = .intPositionIndex
                ' Je regarde l'index courant
                Select Case .intPositionIndex
                  Case Is > vbEmpty
                      ' Je ne suis pas sur le Precedent
                      ' Je n'ai pas lu l'enregistrement
                      .intPositionIndex = .intPositionIndex - 1
                  Case Else
                      ' Je suis sur le premier
                End Select
                ' Je lit le Precedent enregistrement
                Precedent = Lecture()
                ' Je regarde si j'ai bien lu le dernier enregistrement
                Select Case Precedent
                    Case True
                         ' J'ai bien lut l'enregistrement
                    Case False
                         ' Je n'ai pas lu l'enregistrement
                         .intPositionIndex = .intPositionIndex
                         ' Je lit l'ancien enregistrement
                         Lecture
                End Select
        End Select
    End With
   ' Fin
Precedent_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Precedent_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Precedent", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je ne suis pas sur le dernier enregistrement
   Precedent = False
   ' Je Continue
   Resume Precedent_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Suivant() As Boolean                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Se positionner sur le Suivant enregistrement                           ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Suivant - Renvoie si on est sur le                                                    ***
'***                    Suivant enregistrement                                                     ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Suivant()                                                                                       ***
'******************************************************************************
Public Function Suivant() As Boolean
Attribute Suivant.VB_HelpID = 2013
   Dim intPositionIndex As Integer                                                                     'Le numero d'enregistrement courant
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Suivant_Erreur
   ' Je vais
   With mulfiLogFichier.ulinFichierIndex
        ' Je regarde le nombre d'enregistrement
        Select Case .intPositionMax
             Case vbEmpty
                  ' Le nombre d'enregistrement est a zero
                  ' Je renvoie que je ne suis pas sur le Premier enregistrement
                  Suivant = False
             Case Else
                ' Le nombre d'enregistremnt n'est pas a zero
                ' Je me souvient de l'ancien numero d'enregistremnt
                intPositionIndex = .intPositionIndex
                ' Je regarde l'index courant
                Select Case .intPositionIndex
                  Case Is < .intPositionMax
                      ' Je ne suis pas sur le Precedent
                      ' Je n'ai pas lu l'enregistrement
                      .intPositionIndex = .intPositionIndex + 1
                  Case Else
                      ' Je suis sur le premier
                End Select
                ' Je lit le Suivant enregistrement
                Suivant = Lecture()
                ' Je regarde si j'ai bien lu le dernier enregistrement
                Select Case Suivant
                    Case True
                         ' J'ai bien lut l'enregistrement
                    Case False
                         ' Je n'ai pas lu l'enregistrement
                         .intPositionIndex = .intPositionIndex
                         ' Je lit l'ancien enregistrement
                         Lecture
                End Select
        End Select
    End With
   ' Fin
Suivant_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Suivant_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGFONCTION, "Suivant", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie que je ne suis pas sur le dernier enregistrement
   Suivant = False
   ' Je Continue
   Resume Suivant_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Ajouter Une Erreur Dans Le Fichier Journal De L'Application    ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***      objApplication                   - Un Object Application                            ***
'***      intFeuilleType                   - Le Type De La Feuille Source                 ***
'***      strLOGFEUILLENOM                   - Le Nom De La Feuille Source                  ***
'***      intInstructionType            - Le Type De La Instruction           ***
'***      strInstructionNom                       - Le Nom De La Instruction           ***
'***      strInstructionParametre              - Les parametres de la Instruction ***
'***      objErreur                                   - L'Object Erreur                            ***
'***    SORTIE:                                                                                           ***
'***      AjouteErreur                              - Renvoie Si L'Erreur a ete ecrit     ***
'***                                                         dans le Fichier                            ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =AjouteErreur(App,"FORM","FormTestDllAPI","EVENEMENT",         ***
'***        "Form_Load"," ",Err)                                                                     ***
'******************************************************************************
Public Function AjouteErreur(ByVal objApplication As Object, ByVal intFeuilleType As LOGFEUILLETYPETYPEENUM, _
   ByVal strFeuilleNom As String, ByVal IntInstructionType As LOGINSTRUCTIONTYPEENUM, ByVal strInstructionNom As String, _
   ByVal strInstructionParametre As String, ByVal vntErreur As Variant) As Boolean
Attribute AjouteErreur.VB_HelpID = 2014
   Dim lngErreurNumero As Long                                                                  ' Le numero de l'erreur
   Dim strErreurDescription As String                                                          ' La description de l'erreur
   ' Je recupere les donnees de l'erreur avant remise a zero de l'object erreur
   ' Je recupere le numero de l'erreur
   lngErreurNumero = vntErreur.Number
   ' Je recupere la description de l'erreur
   strErreurDescription = vntErreur.Description
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AjouteErreur_Erreur
   ' J'ouvre le fichier
   AjouteErreur = Ouvre(objApplication.Path & "\" & objApplication.EXEName & ".LOG")
   ' Je regarde si j'ai bien ouvert le fichier
   Select Case AjouteErreur
        Case True
            ' Jai bien ouvert le fichier
            ' Je me positionne sur le dernier
            Dernier
            ' Je vais
            With mulfiLogFichier
                ' Je vais
                With .ulinFichierIndex
                    '  Je passe au suivant
                    .intPositionIndex = .intPositionIndex + 1
                    ' Je passe au suivant
                    .intPositionMax = .intPositionMax + 1
                End With
                 ' Je vais beucoup
                 With .uldaFichierDeltaApi
                      ' Je met une chaine vide
                      .strDeltaApiVersion = App.Major & "." & App.Minor & "." & App.Revision
                End With
                ' Je vais
                With .ulapFichierApplication
                    ' Je defini le nom de l'application en erreur
                    .strApplicationNom = objApplication.EXEName
                    ' Je defini la version de l'application en erreur
                    .strApplicationVersion = objApplication.Major & "." & objApplication.Minor & "." & objApplication.Revision
                End With
                ' Je vais
                With .ulfeFichierFeuille
                    ' Je defini le type de la feuille en erreur
                    .intFeuilleType = intFeuilleType
                    ' Je defini le nom de la feuille en erreur
                    .strFeuilleNom = strFeuilleNom
                End With
                ' Je vais
                With .ulinFichierInstruction
                    ' Je defini le type de Instruction en erreur
                    .IntInstructionType = IntInstructionType
                    ' Je defini le nom de la Instruction en erreur
                    .strInstructionNom = strInstructionNom
                    ' Je defini les parametres de l'erreur
                    .strInstructionParametre = strInstructionParametre
                End With
                ' Je vais
                With .ulerFichierErreur
                    ' Je defini le  numero de l'erreur
                    .strErreurNumero = lngErreurNumero
                    ' Je defini la description de l'erreur
                    .strErreurDescription = strErreurDescription
                End With
                ' Je vais
                With .uldtFichierDate
                    ' Je defini la Jour
                    .strDateJour = Date$
                    ' Je defini L'Heure de l'erreur
                    .strDateHeure = Time$
                End With
                ' Je l'AfficheErreur a l'user
                AfficheErreur .ulapFichierApplication.strApplicationNom, .ulapFichierApplication.strApplicationVersion, .ulfeFichierFeuille.intFeuilleType, .ulfeFichierFeuille.strFeuilleNom, _
                   .ulinFichierInstruction.IntInstructionType, .ulinFichierInstruction.strInstructionNom, .ulinFichierInstruction.strInstructionParametre, _
                   .ulerFichierErreur.strErreurNumero, .ulerFichierErreur.strErreurDescription, _
                   .uldtFichierDate.strDateJour, .uldtFichierDate.strDateHeure
                 ' Je l'ecrit
                 AjouteErreur = Ecriture()
                 ' Je regarde si j'ai bien ecrit
                 Select Case AjouteErreur
                      Case True
                           ' Je les bien ecrit
                      Case False
                           ' Je ne l'est pas ecrit
                           ' Je revient au precedent
                           .ulinFichierIndex.intPositionIndex = .ulinFichierIndex.intPositionIndex - 1
                           ' Je lit l'enregistrement
                           Lecture
                 End Select
            End With
        Case False
             ' Je Nai pas ouvert le fichier
   End Select
   ' Je ferme le fichier
   Ferme
   ' Fin
AjouteErreur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
AjouteErreur_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, INSTRUCTIONFONCTION, "AjouteErreur", _
      intFeuilleType & "," & strFeuilleNom & "," & IntInstructionType & "," & strInstructionNom & "," & strInstructionParametre, _
      CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je n'ai pas ajouter l'erreur dans le fichier journal
   AjouteErreur = False
   ' Je Continue
   Resume AjouteErreur_Exit
   ' Fin
End Function

'******************************************************************************
'***    Declaration De Procedure Public                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      Initialise()                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour initialiser les variables                                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      Initialise                                                                                        ***
'******************************************************************************
Public Sub Initialise()
Attribute Initialise.VB_HelpID = 2015
   Dim lngEnrLigneNum As Long
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Initialise_Erreur
    ' Je vais
    With mulfiLogFichier
        ' Je vais beaucoup utiliser les valeurs de l'enregistrement
        With .ulenFichierEnr
             ' Je fait defiler les lignes de l'enregitrements
             For lngEnrLigneNum = LBound(.strEnrLigne) To UBound(.strEnrLigne)
                 ' Je met une chaine vide
                 .strEnrLigne(lngEnrLigneNum) = vbNullString
             Next
         End With
         ' Je vais beucoup
         With .uldaFichierDeltaApi
              ' Je met une chaine vide
              .strDeltaApiVersion = vbNullString
        End With
         ' Je vais beucoup
         With .ulapFichierApplication
              ' Je met une chaine vide
              .strApplicationNom = vbNullString
              ' Je met une chaine vide
              .strApplicationVersion = vbNullString
        End With
         ' Je vais beucoup
         With .ulfeFichierFeuille
              ' Je met une chaine vide
              .intFeuilleType = vbNull
              ' Je met une chaine vide
              .strFeuilleNom = vbNullString
        End With
         ' Je vais beucoup
         With .ulinFichierInstruction
              ' Je met une chaine vide
              .IntInstructionType = vbNull
              ' Je met une chaine vide
              .strInstructionNom = vbNullString
              ' Je met une chaine vide
              .strInstructionParametre = vbNullString
        End With
         ' Je vais beucoup
         With .ulerFichierErreur
              ' Je met une chaine vide
              .strErreurNumero = vbNullString
              ' Je met une chaine vide
              .strErreurDescription = vbNullString
        End With
        ' Je vais beucoup
        With .uldtFichierDate
            ' Je met une chaine vide
            .strDateJour = vbNullString
            ' Je met une chaine vide
            .strDateHeure = vbNullString
        End With
    End With
   ' Fin
Initialise_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
Initialise_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "Initialise", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume Initialise_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      Nouvelle() As Boolean                                                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Se positionner sur le Nouvelle enregistrement                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Nouvelle - Renvoie si on est sur le                                                  ***
'***                    Nouvelle enregistrement                                                   ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =Nouvelle()                                                                                     ***
'******************************************************************************
Public Sub Nouvelle()
Attribute Nouvelle.VB_Description = "En Cas D'Erreur Je Gere L'Erreur"
Attribute Nouvelle.VB_HelpID = 2016
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Nouvelle_Erreur
   ' Je vais
   With mulfiLogFichier.ulinFichierIndex
        ' Je n'ai pas lu l'enregistrement
        .intPositionIndex = .intPositionIndex + 1
        ' Je n'ai pas lu l'enregistrement
        .intPositionMax = .intPositionMax + 1
        ' J'initialise les variables
        Initialise
    End With
   ' Fin
Nouvelle_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
Nouvelle_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "Nouvelle", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume Nouvelle_Exit
   ' Fin
End Sub

'******************************************************************************
'***    Declaration De Procedure Private                                                      ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      LectureDeltaApiVersion                                                                  ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire les donnees de l'application en erreur                                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureDeltaApiVersion()                                                                ***
'******************************************************************************
Private Sub LectureDeltaApiVersion()
   Dim intVersionDebut As Integer                                                               ' Le debut de la chaine version
   Dim intVersionFin As Integer                                                                   ' La fin de la chaine version
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureDeltaApiVersion_Erreur
   ' je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine nom
            intVersionDebut = InStr(1, .strEnrLigne(1), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intVersionFin = InStr(intVersionDebut, .strEnrLigne(1), "]", vbTextCompare)
        End With
        ' Je recupre le nom de l'application
        .ulapFichierApplication.strApplicationNom = Mid$(.ulenFichierEnr.strEnrLigne(1), intVersionDebut, intVersionFin - intVersionDebut)
    End With
   ' Fin
LectureDeltaApiVersion_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureDeltaApiVersion_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureDeltaApiVersion", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureDeltaApiVersion_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      EcritureDeltaApiVersion                                                                  ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire les donnes de l'application en erreur                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureDeltaApiVersion                                                                 ***
'******************************************************************************
Private Sub EcritureDeltaApiVersion()
   Dim strVersion As String                                                                            ' La chaine version
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureDeltaApiVersion_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .uldaFichierDeltaApi
            ' Je defini le nom de l'application en erreur
            strVersion = App.EXEName & " [" & .strDeltaApiVersion & "]" & " Copyright " & App.LegalCopyright
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(1) = "***" & Space$((249 - Len(strVersion)) / 2) & strVersion & Space$((249 - Len(strVersion)) / 2) & "***" & vbCrLf
   End With
   ' Fin
EcritureDeltaApiVersion_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureDeltaApiVersion_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureDeltaApiVersion", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureDeltaApiVersion_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      LectureApplication                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire les donnees de l'application en erreur                                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureApplication()                                                                       ***
'******************************************************************************
Private Sub LectureApplication()
   Dim intNomDebut As Integer                                                                    ' Le debut de la chaine nom
   Dim intNomFin As Integer                                                                        ' La fin de la chaine nom
   Dim intVersionDebut As Integer                                                               ' Le debut de la chaine version
   Dim intVersionFin As Integer                                                                   ' La fin de la chaine version
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureApplication_Erreur
   ' je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine nom
            intNomDebut = InStr(1, .strEnrLigne(3), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intNomFin = InStr(intNomDebut, .strEnrLigne(3), "]", vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intVersionDebut = InStr(intNomFin + 1, .strEnrLigne(3), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intVersionFin = InStrRev(.strEnrLigne(3), "]", 256, vbTextCompare)
        End With
        ' Je recupre le nom de l'application
        .ulapFichierApplication.strApplicationNom = Mid$(.ulenFichierEnr.strEnrLigne(3), intNomDebut, intNomFin - intNomDebut)
        ' Je recupere la version de l'application
        .ulapFichierApplication.strApplicationVersion = Mid$(.ulenFichierEnr.strEnrLigne(3), intVersionDebut, intVersionFin - intVersionDebut)
    End With
   ' Fin
LectureApplication_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureApplication_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureApplication", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureApplication_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      EcritureApplication                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire les donnes de l'application en erreur                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureApplication                                                                         ***
'******************************************************************************
Private Sub EcritureApplication()
   Dim strNom As String                                                                                  ' La chaine nom
   Dim strVersion As String                                                                             ' La chaine version
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureApplication_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulapFichierApplication
            ' Je defini le nom de l'application en erreur
            strNom = LOGAPPLICATION & ":[" & .strApplicationNom & "]"
            ' Je defini la version de l'application en erreur
            strVersion = LOGVERSION & ":[" & .strApplicationVersion & "]"
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(3) = "***" & strNom & Space$(248 - Len(strNom + strVersion)) & strVersion & "***" & vbCrLf
   End With
   ' Fin
EcritureApplication_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureApplication_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureApplication", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureApplication_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      LectureFeuille                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire les donnees de la feuille en erreur                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureFeuille()                                                                              ***
'******************************************************************************
Private Sub LectureFeuille()
   Dim intTypeDebut As Integer                                                                    ' Le debut de la chaine type
   Dim intTypeFin As Integer                                                                        ' La fin de la chaine type
   Dim intNomDebut As Integer                                                                    ' Le debut de la chaine nom
   Dim intNomFin As Integer                                                                        ' La fin de la chaine nom
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureFeuille_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine type
            intTypeDebut = InStr(1, .strEnrLigne(6), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intTypeFin = InStr(intTypeDebut, .strEnrLigne(6), "]", vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intNomDebut = InStr(intTypeFin + 1, .strEnrLigne(6), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intNomFin = InStrRev(.strEnrLigne(6), "]", 256, vbTextCompare)
        End With
        ' Je recupre le type de feuille
        .ulfeFichierFeuille.intFeuilleType = Mid$(.ulenFichierEnr.strEnrLigne(6), intTypeDebut, intTypeFin - intTypeDebut)
        ' Je recupere le nom de la feuille
        .ulfeFichierFeuille.strFeuilleNom = Mid$(.ulenFichierEnr.strEnrLigne(6), intNomDebut, intNomFin - intNomDebut)
    End With
   ' Fin
LectureFeuille_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureFeuille_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureFeuille", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureFeuille_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      EcritureFeuille                                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Ecrire les donnees de la feuille en erreur                                 ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureFeuille()                                                                              ***
'******************************************************************************
Private Sub EcritureFeuille()
   Dim strType As String                                                                                 ' La chaine type
   Dim strNom As String                                                                                  ' La chaine nom
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureFeuille_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulfeFichierFeuille
            ' Je defini le type  de la feuille en erreur
            strType = LOGTYPE & ":[" & .intFeuilleType & "]"
            ' Je defini le nom  de la feuille en erreur
            strNom = LOGNOM & ":[" & .strFeuilleNom & "]"
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(6) = "***" & strType & Space$(248 - Len(strType + strNom)) & strNom & "***" & vbCrLf
    End With
   ' Fin
EcritureFeuille_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureFeuille_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureFeuille", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureFeuille_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      LectureInstruction                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire les donnees de la Instruction en erreur                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureInstruction()                                                                        ***
'******************************************************************************
Private Sub LectureInstruction()
   Dim intTypeDebut As Integer                                                                    ' Le debut de la chaine type
   Dim intTypeFin As Integer                                                                        ' La fin de la chaine type
   Dim intNomDebut As Integer                                                                    ' Le debut de la chaine nom
   Dim intNomFin As Integer                                                                        ' La fin de la chaine nom
   Dim intParametreADebut As Integer                                                         ' Le debut de la chaine parametre
   Dim intParametreAFin As Integer                                                             ' La fin de la chaine parametre
   Dim intParametreBDebut As Integer                                                        ' Le debut de la chaine parametre
   Dim intParametreBFin As Integer                                                            ' La fin de la chaine parametre
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureInstruction_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine type
            intTypeDebut = InStr(1, .strEnrLigne(9), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine type
            intTypeFin = InStr(intTypeDebut, .strEnrLigne(9), "]", vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intNomDebut = InStr(intTypeFin + 1, .strEnrLigne(9), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intNomFin = InStrRev(.strEnrLigne(9), "]", 256, vbTextCompare)
            ' Je calcule le debut de la chaine parametre
            intParametreADebut = InStr(1, .strEnrLigne(11), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine parametre
            intParametreAFin = InStrRev(.strEnrLigne(11), "]", 256, vbTextCompare)
            ' Je calcule le debut de la chaine parametre
            intParametreBDebut = InStr(1, .strEnrLigne(12), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine parametre
            intParametreBFin = InStrRev(.strEnrLigne(12), "]", 256, vbTextCompare)
        End With
        ' Je recupre le type d'instruction
        .ulinFichierInstruction.IntInstructionType = Mid$(.ulenFichierEnr.strEnrLigne(9), intTypeDebut, intTypeFin - intTypeDebut)
        ' Je recupere le nom de l'instruction
        .ulinFichierInstruction.strInstructionNom = Mid$(.ulenFichierEnr.strEnrLigne(9), intNomDebut, intNomFin - intNomDebut)
        ' Je recupere les parametres de l'instruction
        .ulinFichierInstruction.strInstructionParametre = Mid$(.ulenFichierEnr.strEnrLigne(11), intParametreADebut, intParametreAFin - intParametreADebut) & Mid$(.ulenFichierEnr.strEnrLigne(12), intParametreBDebut, intParametreBFin - intParametreBDebut)
    End With
   ' Fin
LectureInstruction_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureInstruction_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureInstruction", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureInstruction_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      EcritureInstruction                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Ecrire les donnees de la Instruction en erreur                          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureInstruction()                                                                       ***
'******************************************************************************
Private Sub EcritureInstruction()
   Dim strType As String                                                                                 ' La chaine type
   Dim strNom As String                                                                                  ' La chaine nom
   Dim strParametreA As String                                                                       ' La chaine parametre A
   Dim lngParametreALongueur As Long                                                           ' La longueur de la chaine parametre A
   Dim strParametreB As String                                                                       ' La chaine parametre B
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureInstruction_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulinFichierInstruction
            ' Je defini le type de l'instruction en erreur
            strType = LOGTYPE & ":[" & .IntInstructionType & "]"
            ' Je defini le nom de l'instruction en erreur
            strNom = LOGNOM & ":[" & .strInstructionNom & "]"
            ' Je regare la longueur de la chaine parametre
            Select Case Len(.strInstructionParametre)
                Case Is > 128
                    ' La chaine est plus grande que 68
                    ' Je cherche une virgule
                    lngParametreALongueur = InStrRev(.strInstructionParametre, ",", 128)
                    ' Je regarde la longueur
                    Select Case lngParametreALongueur
                        Case vbEmpty
                            '  La longueur est null
                            ' Je la defini
                            lngParametreALongueur = 128
                        Case Else
                            '  La longueur n'est pas null
                            ' Je la defini
                            lngParametreALongueur = lngParametreALongueur Mod 128
                    End Select
                    ' Je defini le descriptif  de l'erreur
                    strParametreA = "[" & Left(.strInstructionParametre, lngParametreALongueur) & "]"
                    ' Je defini le descriptif  de l'erreur
                    strParametreB = "[" & Mid(.strInstructionParametre, lngParametreALongueur, 128) & "]"
                Case Else
                    ' Je defini le descriptif  de l'erreur
                    strParametreA = "[" & .strInstructionParametre & "]"
                    ' Je defini le descriptif  de l'erreur
                    strParametreB = "[" & "]"
            End Select
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(9) = "***" & strType & Space$(248 - Len(strType + strNom)) & strNom & "***" & vbCrLf
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(11) = "***" & strParametreA & Space$(248 - Len(strParametreA)) & "***" & vbCrLf
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(12) = "***" & strParametreB & Space$(248 - Len(strParametreB)) & "***" & vbCrLf
    End With
   ' Fin
EcritureInstruction_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureInstruction_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureInstruction", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureInstruction_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      LectureErreur                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lire les donnees de l'erreur                                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureErreur                                                                                 ***
'******************************************************************************
Private Sub LectureErreur()
   Dim intNumeroDebut As Integer                                                               ' Le debut de la chaine numero
   Dim intNumeroFin As Integer                                                                   ' La fin de la chaine numero
   Dim intDescriptionADebut As Integer                                                         ' Le debut de la chaine description
   Dim intDescriptionAFin As Integer                                                             ' La fin de la chaine description
   Dim intDescriptionBDebut As Integer                                                         ' Le debut de la chaine description
   Dim intDescriptionBFin As Integer                                                             ' La fin de la chaine description
   Dim intDescriptionCDebut As Integer                                                         ' Le debut de la chaine description
   Dim intDescriptionCFin As Integer                                                             ' La fin de la chaine description
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureErreur_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine numero
            intNumeroDebut = InStr(1, .strEnrLigne(14), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine numero
            intNumeroFin = InStr(intNumeroDebut, .strEnrLigne(14), "]", vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intDescriptionADebut = InStr(1, .strEnrLigne(16), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intDescriptionAFin = InStrRev(.strEnrLigne(16), "]", 256, vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intDescriptionBDebut = InStr(1, .strEnrLigne(17), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intDescriptionBFin = InStrRev(.strEnrLigne(17), "]", 256, vbTextCompare)
            ' Je calcule le debut de la chaine nom
            intDescriptionCDebut = InStr(1, .strEnrLigne(18), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine nom
            intDescriptionCFin = InStrRev(.strEnrLigne(18), "]", 256, vbTextCompare)
        End With
        ' Je recupre le numero de l'erreur
        .ulerFichierErreur.strErreurNumero = Mid$(.ulenFichierEnr.strEnrLigne(14), intNumeroDebut, intNumeroFin - intNumeroDebut)
        ' Je recupere la description de l'erreur
        .ulerFichierErreur.strErreurDescription = Mid$(.ulenFichierEnr.strEnrLigne(16), intDescriptionADebut, intDescriptionAFin - intDescriptionADebut) _
            & Mid$(.ulenFichierEnr.strEnrLigne(17), intDescriptionBDebut, intDescriptionBFin - intDescriptionBDebut) _
            & Mid$(.ulenFichierEnr.strEnrLigne(18), intDescriptionCDebut, intDescriptionCFin - intDescriptionCDebut)
    End With
   ' Fin
LectureErreur_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureErreur_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureErreur", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureErreur_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***      EcritureErreur                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire les donnes  de l'erreur                                                    ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureErreur                                                                                ***
'******************************************************************************
Private Sub EcritureErreur()
   Dim strNumero As String                                                                             ' La chaine numero
   Dim strDescriptionA As String                                                                     ' La chaine description A
   Dim lngDescriptionALongueur As Long                                                         ' La longueur de la chaine description A
   Dim strDescriptionB As String                                                                     ' La chaine description B
   Dim lngDescriptionBLongueur As Long                                                         ' La longueur de la chaine description B
   Dim strDescriptionC As String                                                                     ' La chaine description C
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureErreur_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulerFichierErreur
            ' Je defini le numero de l'erreur
            strNumero = LOGNUMERO & ":[" & .strErreurNumero & "]"
            ' Je regare la longueur de la chaine parametre
            Select Case Len(.strErreurDescription)
                Case Is > 128
                    ' La chaine est plus grande que 68
                    ' Je cherche une virgule
                    lngDescriptionALongueur = InStrRev(.strErreurDescription, ".", 128)
                    ' Je regarde la longueur
                    Select Case lngDescriptionALongueur
                        Case vbEmpty
                            '  La longueur est null
                            ' Je cherche une virgule
                            lngDescriptionALongueur = InStrRev(.strErreurDescription, " ", 128)
                            ' Je regarde la longueur
                            Select Case lngDescriptionALongueur
                                Case vbEmpty
                                    '  La longueur est null
                                    ' Je la defini
                                    lngDescriptionALongueur = 128
                                Case Else
                                    '  La longueur n'est pas null
                                    ' Je la defini
                                    lngDescriptionALongueur = lngDescriptionALongueur Mod 128
                            End Select
                        Case Else
                            '  La longueur n'est pas null
                            ' Je la defini
                            lngDescriptionALongueur = lngDescriptionALongueur Mod 128
                    End Select
                    ' Je defini le descriptif  de l'erreur
                    strDescriptionA = "[" & Left(.strErreurDescription, lngDescriptionALongueur) & "]"
                    ' Je regare la longueur de la chaine parametre
                    Select Case Len(.strErreurDescription)
                        Case Is > 256
                            ' La chaine est plus grande que 256
                            ' Je cherche une virgule
                            lngDescriptionBLongueur = InStrRev(.strErreurDescription, ".", 256) - 128
                            ' Je regarde la longueur
                            Select Case lngDescriptionBLongueur
                                Case Is < vbEmpty
                                    '  La longueur est null
                                    ' Je cherche une virgule
                                    lngDescriptionBLongueur = InStrRev(.strErreurDescription, " ", 256) - 128
                                    ' Je regarde la longueur
                                    Select Case lngDescriptionBLongueur
                                        Case Is < vbEmpty
                                            '  La longueur est null
                                            ' Je la defini
                                            lngDescriptionBLongueur = 128
                                        Case Else
                                            '  La longueur n'est pas null
                                            ' Je la defini
                                            lngDescriptionBLongueur = lngDescriptionBLongueur Mod 128
                                    End Select
                                Case Else
                                    '  La longueur n'est pas null
                                    ' Je la defini
                                    lngDescriptionBLongueur = lngDescriptionBLongueur Mod 128
                            End Select
                            ' Je defini le descriptif  de l'erreur
                            strDescriptionB = "[" & Mid(.strErreurDescription, lngDescriptionALongueur, lngDescriptionBLongueur) & "]"
                            ' Je defini le descriptif  de l'erreur
                            strDescriptionC = "[" & Mid(.strErreurDescription, lngDescriptionALongueur + lngDescriptionBLongueur, 128) & "]"
                        Case Else
                            ' La chaine n'est pas plus grand que 136
                            ' Je defini le descriptif  de l'erreur
                            strDescriptionB = "[" & Mid(.strErreurDescription, lngDescriptionALongueur) & "]"
                            ' Je defini le descriptif  de l'erreur
                            strDescriptionC = "[" & "]"
                    End Select
                Case Else
                    ' La chaine n'est pas plus grand que 68
                    ' Je defini le descriptif  de l'erreur
                    strDescriptionA = "[" & .strErreurDescription & "]"
                    ' Je defini le descriptif  de l'erreur
                    strDescriptionB = "[" & "]"
                    ' Je defini le descriptif  de l'erreur
                    strDescriptionC = "[" & "]"
            End Select
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(14) = "***" & strNumero & Space$(248 - Len(strNumero)) & "***" & vbCrLf
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(16) = "***" & strDescriptionA & Space$(248 - Len(strDescriptionA)) & "***" & vbCrLf
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(17) = "***" & strDescriptionB & Space$(248 - Len(strDescriptionB)) & "***" & vbCrLf
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(18) = "***" & strDescriptionC & Space$(248 - Len(strDescriptionC)) & "***" & vbCrLf
    End With
   ' Fin
EcritureErreur_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureErreur_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureErreur", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureErreur_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      LectureDate                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire les donnnes de la Date de l'erreur                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LectureDate                                                                                    ***
'******************************************************************************
Private Sub LectureDate()
   Dim intJourDebut As Integer                                                                     ' Le debut de la chaine jour
   Dim intJourFin As Integer                                                                         ' La fin de la chaine jour
   Dim intHeureDebut As Integer                                                                   ' Le debut de la chaine heure
   Dim intHeureFin As Integer                                                                       ' La fin de la chaine heure
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureDate_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .ulenFichierEnr
            ' Je calcule le debut de la chaine jour
            intJourDebut = InStr(1, .strEnrLigne(20), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine jour
            intJourFin = InStr(intJourDebut, .strEnrLigne(20), "]", vbTextCompare)
            ' Je calcule le debut de la chaine heure
            intHeureDebut = InStr(intJourFin, .strEnrLigne(20), "[", vbTextCompare) + 1
            ' Je cacule la fin de la chaine heure
            intHeureFin = InStrRev(.strEnrLigne(20), "]", 256, vbTextCompare)
        End With
        ' Je recupere le jour de l'erreur
        .uldtFichierDate.strDateJour = Mid$(.ulenFichierEnr.strEnrLigne(20), intJourDebut, intJourFin - intJourDebut)
        ' Je recupere l'heure de l'erreur
        .uldtFichierDate.strDateHeure = Mid$(.ulenFichierEnr.strEnrLigne(20), intHeureDebut, intHeureFin - intHeureDebut)
    End With
   ' Fin
LectureDate_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
LectureDate_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "LectureDate", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume LectureDate_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      EcritureDate                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire les donnes de la date de l'erreur                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      EcritureDate                                                                                   ***
'******************************************************************************
Private Sub EcritureDate()
   Dim strJour As String                                                                                ' La chaine jour
   Dim strHeure As String                                                                              ' La chaine heure
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureDate_Erreur
   ' Je vais
   With mulfiLogFichier
        ' Je vais
        With .uldtFichierDate
            ' Je defini le Jour de l'erreur
            strJour = LOGJOUR & ":[" & .strDateJour & "]"
            ' Je defini l'heure de l'erreur
            strHeure = LOGHEURE & ":[" & .strDateHeure & "]"
        End With
        ' Je defini la fin de la ligne
        .ulenFichierEnr.strEnrLigne(20) = "***" & strJour & Space$(248 - Len(strJour & strHeure)) & strHeure & "***" & vbCrLf
    End With
   ' Fin
EcritureDate_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
EcritureDate_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "EcritureDate", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume EcritureDate_Exit
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***     AfficheErreur(ByVal strApplicationNom As String,                            ***
'***     ByVal strApplicationVersion As String,                                            ***
'***     ByVal strFeuilleType As String, ByVal strLOGFEUILLENOM As String,       ***
'***     ByVal strInstructionType As String,                                                 ***
'***     ByVal strInstructionNom As String,                                                  ***
'***     ByVal strInstructionParametre As String,                                         ***
'***     ByVal strErreurNumero As String,                                                    ***
'***     ByVal strErreurDescription As String,                                              ***
'***     ByVal strDateJour As String, ByVal strDateHeure As String)            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour AfficheErreurr les donnes de l'erreur a l'user                             ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      AfficheErreur App.EXEName, vbnullstring, conFEUILLECLASS,              ***
'***      conLOGFEUILLENOM,conLOGPROCEDURE, "EcritureDate", vbnullstring,    ***
'***      CStr(Err.Number),                                                                           ***
'***      Err.Description, vbnullstring, vbnullstring                                       ***
'******************************************************************************
Private Sub AfficheErreur(ByVal strApplicationNom As String, ByVal strApplicationVersion As String, ByVal intFeuilleType As Integer, _
   ByVal strLOGFEUILLENOM As String, ByVal IntInstructionType As Integer, ByVal strInstructionNom As String, _
   ByVal strInstructionParametre As String, ByVal strErreurNumero As String, ByVal strErreurDescription As String, _
   ByVal strDateJour As String, ByVal strDateHeure As String)
   Dim strErreurMessage(12) As String * 40
   ' Je Gere pas l'erreur
   On Error Resume Next
   ' Je defini la premiere ligne du message d'erreur
   strErreurMessage(0) = LOGAPPLICATION & ":"
   ' Je defini la premiere ligne du message d'erreur
   strErreurMessage(0) = strErreurMessage(0) & String(10 - Len(strErreurMessage(0)), " ") & LOGFEUILLE & ":" & vbCrLf
   ' Je defini la deuxieme ligne du message d'erreur
   strErreurMessage(1) = LOGNOM & ":" & strApplicationNom
   ' Je defini la deuxieme ligne du message d'erreur
   strErreurMessage(1) = strErreurMessage(1) & String(10 - Len(strErreurMessage(1)), " ") & LOGTYPE & ":" & intFeuilleType & vbCrLf
   ' Je defini la troisieme ligne du message d'erreur
   strErreurMessage(2) = LOGVERSION & ":" & strApplicationVersion
   ' Je defini la troisieme ligne du message d'erreur
   strErreurMessage(2) = strErreurMessage(2) & String(10 - Len(strErreurMessage(2)), " ") & LOGNOM & ":" & strLOGFEUILLENOM & vbCrLf & vbCrLf
   ' Je defini la quatrieme ligne du message d'erreur
   strErreurMessage(3) = LOGINSTRUCTION & ":" & vbCrLf
   ' Je defini la cinquieme ligne du message d'erreur
   strErreurMessage(4) = LOGTYPE & ":" & IntInstructionType & vbCrLf
   ' Je defini la sixieme ligne du message d'erreur
   strErreurMessage(5) = LOGNOM & ":" & strInstructionNom & vbCrLf
   ' Je defini la septieme ligne du message d'erreur
   strErreurMessage(6) = LOGPARAMETRE & ":" & strInstructionParametre & vbCrLf & vbCrLf
   ' Je defini la huitieme ligne du message d'erreur
   strErreurMessage(7) = LOGERREUR & ":" & vbCrLf
   ' Je defini la neuvieme ligne du message d'erreur
   strErreurMessage(8) = LOGNUMERO & ":" & strErreurNumero & vbCrLf
   ' Je defini la dixieme ligne du message d'erreur
   strErreurMessage(9) = LOGDESCRIPTION & ":" & strErreurDescription & vbCrLf & vbCrLf
   ' Je defini la onzieme ligne du message d'erreur
   strErreurMessage(10) = LOGJOUR & ":" & strDateJour & vbCrLf
   ' Je defini la douzieme ligne du message d'erreur
   strErreurMessage(11) = LOGHEURE & ":" & strDateHeure & vbCrLf
   ' J'AfficheErreur L'Erreur A L'user
   MsgBox strErreurMessage(0) & strErreurMessage(1) & strErreurMessage(2) & strErreurMessage(3) & strErreurMessage(4) & _
             strErreurMessage(5) & strErreurMessage(6) & strErreurMessage(7) & strErreurMessage(8) & strErreurMessage(9) & _
             strErreurMessage(10) & strErreurMessage(11) _
             , , strApplicationNom & " " & strApplicationVersion & " " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
   ' Fin
AfficheErreur_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
End Sub

'******************************************************************************
'***    Declaration De Propriete Public                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DeltaApiVersion                                                                             ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner la version de la dll qui a ecrit les informations          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      DeltaApiVersion  - Le version de la dll                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =ApplicationNom                                                                            ***
'******************************************************************************
Public Property Get DeltaApiVersion() As String
Attribute DeltaApiVersion.VB_HelpID = 2017
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo DeltaApiVersion_Erreur
   ' Je renvoie la version de la dll
   DeltaApiVersion = mulfiLogFichier.uldaFichierDeltaApi.strDeltaApiVersion
   ' Fin
DeltaApiVersion_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
DeltaApiVersion_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DeltaApiVersion", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   DeltaApiVersion = vbNullString
   ' Je Continue
   Resume DeltaApiVersion_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DeltaApiVersion                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir la version de la dll qui a ecrit les informations               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strDeltaApiVersion - la version de la dll qui a ecrit les informations***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      DeltaApiVersion="1.0.01"                                                                ***
'******************************************************************************
Public Property Let DeltaApiVersion(ByVal strDeltaApiVersion As String)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo DeltaApiVersion_Erreur
   'Je defini la version de la dll
   mulfiLogFichier.ulapFichierApplication.strApplicationNom = strDeltaApiVersion
   ' Fin
DeltaApiVersion_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
DeltaApiVersion_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DeltaApiVersion", strDeltaApiVersion, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume DeltaApiVersion_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ApplicationNom                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le nom de l'application en erreur                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      ApplicationNom  - Le nom de l'application en erreur                         ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =ApplicationNom                                                                            ***
'******************************************************************************
Public Property Get ApplicationNom() As String
Attribute ApplicationNom.VB_HelpID = 2018
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo ApplicationNom_Erreur
   ' Je renvoie le nom de l'application
   ApplicationNom = mulfiLogFichier.ulapFichierApplication.strApplicationNom
   ' Fin
ApplicationNom_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
ApplicationNom_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ApplicationNom", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   ApplicationNom = vbNullString
   ' Je Continue
   Resume ApplicationNom_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ApplicationNom                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le nom de l'application en erreur                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strApplicationNom - Le nom de l'application en erreur                    ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      ApplicationNom="Projet"                                                                ***
'******************************************************************************
Public Property Let ApplicationNom(ByVal strApplicationNom As String)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo ApplicationNom_Erreur
   'Je defini le nom de l'application en erreur
   mulfiLogFichier.ulapFichierApplication.strApplicationNom = strApplicationNom
   ' Fin
ApplicationNom_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
ApplicationNom_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ApplicationNom", strApplicationNom, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume ApplicationNom_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ApplicationVersion                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner la version de l'application en erreur                          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      ApplicationVersion  - La version de l'application en erreur               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =ApplicationVersion                                                                       ***
'******************************************************************************
Public Property Get ApplicationVersion() As String
Attribute ApplicationVersion.VB_HelpID = 2019
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo ApplicationVersion_Erreur
   ' Je renvoie la version de l'application
   ApplicationVersion = mulfiLogFichier.ulapFichierApplication.strApplicationVersion
   ' Fin
ApplicationVersion_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
ApplicationVersion_Erreur:
   ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ApplicationVersion", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   ApplicationVersion = vbNullString
   ' Je Continue
   Resume ApplicationVersion_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ApplicationVersion                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir la version de l'application en erreur                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strApplicationVersion - La version de l'application en erreur          ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      ApplicationVersion="Projet"                                                           ***
'******************************************************************************
Public Property Let ApplicationVersion(ByVal strApplicationVersion As String)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo ApplicationVersion_Erreur
   'Je defini la version de l'application en erreur
   mulfiLogFichier.ulapFichierApplication.strApplicationVersion = strApplicationVersion
   ' Fin
ApplicationVersion_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
ApplicationVersion_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ApplicationVersion", strApplicationVersion, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume ApplicationVersion_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      FeuilleType                                                                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le type de feuille en l'erreur                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      FeuilleType  - le type de feuille en l'erreur                                       ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =FeuilleType                                                                                   ***
'******************************************************************************
Public Property Get FeuilleType() As LOGFEUILLETYPETYPEENUM
Attribute FeuilleType.VB_HelpID = 2020
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo FeuilleType_Erreur
    ' Je retourne le type de feuille
    FeuilleType = mulfiLogFichier.ulfeFichierFeuille.intFeuilleType
    ' Fin
FeuilleType_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
FeuilleType_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "FeuilleType", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   FeuilleType = vbNullString
    ' Je Continue
    Resume FeuilleType_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      FeuilleType                                                                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le type de feuille en erreur                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strFeuilleType - Le type de feuille en erreur                                   ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      FeuilleType="FORM:"                                                                      ***
'******************************************************************************
Public Property Let FeuilleType(ByVal intFeuilleType As LOGFEUILLETYPETYPEENUM)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo FeuilleType_Erreur
    ' Je defini la feuille en erreur
    mulfiLogFichier.ulfeFichierFeuille.intFeuilleType = intFeuilleType
    ' Fin
FeuilleType_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
FeuilleType_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "FeuilleType", intFeuilleType, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume FeuilleType_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      LOGFEUILLENOM                                                                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le nom de feuille en l'erreur                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      LOGFEUILLENOM  - le nom de feuille en l'erreur                                        ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =LOGFEUILLENOM                                                                                   ***
'******************************************************************************
Public Property Get LOGFEUILLENOM() As String
Attribute LOGFEUILLENOM.VB_HelpID = 2021
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo LOGFEUILLENOM_Erreur
    ' Je retourne le nom de feuille
    LOGFEUILLENOM = mulfiLogFichier.ulfeFichierFeuille.strFeuilleNom
    ' Fin
LOGFEUILLENOM_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
LOGFEUILLENOM_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "LOGFEUILLENOM", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   LOGFEUILLENOM = vbNullString
    ' Je Continue
    Resume LOGFEUILLENOM_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      LOGFEUILLENOM                                                                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le nom de feuille en erreur                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      LOGFEUILLENOM _Str - Le Nom  de feuille en erreur                                 ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      LOGFEUILLENOM ="FORM:"                                                                     ***
'******************************************************************************
Public Property Let LOGFEUILLENOM(ByVal strLOGFEUILLENOM As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo LOGFEUILLENOM_Erreur
    ' Je defini la feuille en erreur
    mulfiLogFichier.ulfeFichierFeuille.strFeuilleNom = strLOGFEUILLENOM
    ' Fin
LOGFEUILLENOM_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
LOGFEUILLENOM_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "LOGFEUILLENOM", strLOGFEUILLENOM, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume LOGFEUILLENOM_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionType                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le type de Instruction en l'erreur                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      InstructionType  - le type de Instruction en l'erreur                         ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =InstructionType                                                                            ***
'******************************************************************************
Public Property Get InstructionType() As Integer
Attribute InstructionType.VB_HelpID = 2022
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionType_Erreur
    ' Je retourne le type de Instruction
    InstructionType = mulfiLogFichier.ulinFichierInstruction.IntInstructionType
    ' Fin
InstructionType_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionType_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionType", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   InstructionType = vbNullString
    ' Je Continue
    Resume InstructionType_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionType                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le type de Instruction en erreur                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strInstructionType - Le type de Instruction en erreur                      ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      InstructionType="PROCEDURE:"                                                    ***
'******************************************************************************
Public Property Let InstructionType(ByVal IntInstructionType As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionType_Erreur
    ' Je defini le type de Instruction en erreur
    mulfiLogFichier.ulinFichierInstruction.IntInstructionType = IntInstructionType
    ' Fin
InstructionType_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionType_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionType", IntInstructionType, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume InstructionType_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionNom                                                                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le nom de Instruction en l'erreur                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      InstructionNom  - le nom de Instruction en l'erreur                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =InstructionNom                                                                             ***
'******************************************************************************
Public Property Get InstructionNom() As String
Attribute InstructionNom.VB_HelpID = 2023
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionNom_Erreur
    ' Je retourne le nom de la Instruction  en erreur
    InstructionNom = mulfiLogFichier.ulinFichierInstruction.strInstructionNom
    ' Fin
InstructionNom_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionNom_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionNom", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   InstructionNom = vbNullString
    ' Je Continue
    Resume InstructionNom_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionNom                                                                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le nom de la Instruction en erreur                                  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strInstructionNom - Le Nom de la Instruction en erreur                  ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      InstructionNom ="ErreurLOGFEUILLENOM"                                               ***
'******************************************************************************
Public Property Let InstructionNom(ByVal strInstructionNom As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionNom_Erreur
    ' Je defini la feuille en erreur
    mulfiLogFichier.ulinFichierInstruction.strInstructionNom = strInstructionNom
    ' Fin
InstructionNom_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionNom_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionNom", strInstructionNom, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume InstructionNom_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionParametre                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner les parametres de la Instruction en l'erreur               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      InstructionParametre  - les parametres de la Instruction en l'erreur  ***
'******************************************************************************
'***    EXEMPLE:                                                                                         ***
'***      =InstructionParametre                                                                     ***
'******************************************************************************
Public Property Get InstructionParametre() As String
Attribute InstructionParametre.VB_HelpID = 2024
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionParametre_Erreur
    ' Je retourne le nom de la Instruction  en erreur
    InstructionParametre = mulfiLogFichier.ulinFichierInstruction.strInstructionParametre
    ' Fin
InstructionParametre_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionParametre_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionParametre", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   InstructionParametre = vbNullString
    ' Je Continue
    Resume InstructionParametre_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      InstructionParametre                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir les parametres de la Instruction en erreur                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strInstructionParametre - les parametre de la Instruction en erreur***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      InstructionParametre ="TI,TO"                                                        ***
'******************************************************************************
Public Property Let InstructionParametre(ByVal strInstructionParametre As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InstructionParametre_Erreur
    ' Je defini la feuille en erreur
    mulfiLogFichier.ulinFichierInstruction.strInstructionParametre = strInstructionParametre
    ' Fin
InstructionParametre_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
InstructionParametre_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "InstructionParametre", strInstructionParametre, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume InstructionParametre_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ErreurNumero                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le Numero de l'erreur                                                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      ErreurNumero  - le Numero de l'erreur                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =ErreurNumero                                                                               ***
'******************************************************************************
Public Property Get ErreurNumero() As String
Attribute ErreurNumero.VB_HelpID = 2025
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ErreurNumero_Erreur
    ' Je renvoie le numero
    ErreurNumero = mulfiLogFichier.ulerFichierErreur.strErreurNumero
    ' Fin
ErreurNumero_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
ErreurNumero_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ErreurNumero", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   ErreurNumero = vbNullString
    ' Je Continue
    Resume ErreurNumero_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ErreurNumero                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir le Numero de l'erreur                                                    ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strErreurNumero - La Numero de l'erreur                                        ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      ErreurNumero="405"                                                                       ***
'******************************************************************************
Public Property Let ErreurNumero(ByVal strErreurNumero As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ErreurNumero_Erreur
    ' Je defini le numero en erreur
    mulfiLogFichier.ulerFichierErreur.strErreurNumero = strErreurNumero
    ' Fin
ErreurNumero_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
ErreurNumero_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ErreurNumero", strErreurNumero, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume ErreurNumero_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ErreurDescription                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner la Description de l'erreur                                          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      ErreurDescription  - la Description de l'erreur                                  ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =ErreurDescription                                                                         ***
'******************************************************************************
Public Property Get ErreurDescription() As String
Attribute ErreurDescription.VB_HelpID = 2026
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ErreurDescription_Erreur
    ' Je renvoie la Jour
    ErreurDescription = mulfiLogFichier.ulerFichierErreur.strErreurDescription
    ' Fin
ErreurDescription_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
ErreurDescription_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ErreurDescription", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   ErreurDescription = vbNullString
    ' Je Continue
    Resume ErreurDescription_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      ErreurDescription                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir la description de l'erreur                                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strErreurDescription - La description de l'erreur                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      ErreurDescription="Type Incompatible"                                           ***
'******************************************************************************
Public Property Let ErreurDescription(ByVal strErreurDescription As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ErreurDescription_Erreur
    ' Je defini le nom de l'application en erreur
    mulfiLogFichier.ulerFichierErreur.strErreurDescription = strErreurDescription
    ' Fin
ErreurDescription_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
ErreurDescription_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
    AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "ErreurDescription", strErreurDescription, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume ErreurDescription_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DateJour                                                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le Jour de l'erreur                                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      DateJour  - le Jour de l'erreur                                                         ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =DateJour                                                                                      ***
'******************************************************************************
Public Property Get DateJour() As String
Attribute DateJour.VB_HelpID = 2027
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo DateJour_Erreur
    ' Je renvoie la Jour
    DateJour = mulfiLogFichier.uldtFichierDate.strDateJour
    ' Fin
DateJour_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
DateJour_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DateJour", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   DateJour = vbNullString
    ' Je Continue
    Resume DateJour_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DateJour                                                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir la Jour de l'erreur                                                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strDateJour - La Jour de l'erreur                                                    ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      DateJour="10/12/2000"                                                                    ***
'******************************************************************************
Public Property Let DateJour(ByVal strDateJour As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo DateJour_Erreur
    ' Je defini le nom de l'application en erreur
    mulfiLogFichier.uldtFichierDate.strDateJour = strDateJour
    ' Fin
DateJour_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
DateJour_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DateJour", strDateJour, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume DateJour_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DateHeure                                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour retourner le Heure de l'erreur                                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      DateHeure  - l'Heure de l'erreur                                                       ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =DateHeure                                                                                    ***
'******************************************************************************
Public Property Get DateHeure() As String
Attribute DateHeure.VB_HelpID = 2028
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo DateHeure_Erreur
    ' Je renvoie l'heure
    DateHeure = mulfiLogFichier.uldtFichierDate.strDateHeure
    ' Fin
DateHeure_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
DateHeure_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DateHeure", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je renvoie une chaine vide
   DateHeure = vbNullString
    ' Je Continue
    Resume DateHeure_Exit
    ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***      DateHeure                                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour definir l'heure de l'erreur                                                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strDateHeure - L'heure de l'erreur                                                  ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      DateHeure="10/12/2000"                                                                 ***
'******************************************************************************
Public Property Let DateHeure(ByVal strDateHeure As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo DateHeure_Erreur
    ' Je defini le nom de l'application en erreur
    mulfiLogFichier.uldtFichierDate.strDateHeure = strDateHeure
    ' Fin
DateHeure_Exit:
    ' Je Sort De La Procedure
    Exit Property
    ' Fin
DateHeure_Erreur:
    ' J'AfficheErreur L'Erreur A L'user
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROPRIETE, "DateHeure", strDateHeure, CStr(Err.Number), Err.Description, vbNullString, vbNullString
    ' Je Continue
    Resume DateHeure_Exit
    ' Fin
End Property

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
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "Class_Initialize", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
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
   AfficheErreur App.EXEName, vbNullString, FEUILLECLASS, LOGFEUILLENOM, LOGPROCEDURE, "Class_Terminate", vbNullString, CStr(Err.Number), Err.Description, vbNullString, vbNullString
   ' Je Continue
   Resume Class_Terminate_Exit
   ' Fin
End Sub

