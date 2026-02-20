VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRegWinFicApl"
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
'***      Pour La Gestion Des Command De Fichier Et Command Application  ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Ajout,Suppression,Recherche,Recuperation et Enumeration              ***
'***      des command de fichier et command application                              ***
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
Private Const mconFEUILLETYPE = "CLASS"                                                      ' La feuille du source
Private Const mconFEUILLENOM = "clsWinRegFicApl"                                        ' Je me rapelle le nom de la Feuille

'******************************************************************************
'***    Constante Qui Defini L'emplacement de la class de fichier                 ***
'******************************************************************************
Private Const mconrcrClassFichier = HKEY_CLASSES_ROOT
Private Const mconrcsClassFichier = ""

'******************************************************************************
'***    Constante Qui Defini L'emplacement de la class application               ***
'******************************************************************************
Private Const mconrcrClassApplication = HKEY_CLASSES_ROOT
Private Const mconrcsClassApplication = ""

'******************************************************************************
'***    Constante Qui Defini L'emplacement des commandes                        ***
'******************************************************************************
Private Const mconrcsIcon = "DefaultIcon"

'******************************************************************************
'***    Constante Qui Defini L'emplacement des commandes                        ***
'******************************************************************************
Private Const mconrcsCommand = "Command"

'******************************************************************************
'***    Constante Qui Defini L'emplacement des commandes                        ***
'******************************************************************************
Private Const mconrcsBureau = "Shell"
Private Const mconrcsBureauNouveau = "ShellNew"

'******************************************************************************
'***    Constante Qui Defini L'emplacement                                                 ***
'******************************************************************************
Private Const mconrcsOuvre = "Open"
Private Const mconrcsImprime = "Print"
Private Const mconrcsImprimeVers = "PrintTo"

'******************************************************************************
'***    Declaration De Type Public                                                              ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion De La Class Fichiers                                        ***
'******************************************************************************
Public Type CLSFIC
   strClsFicExtension As String
   strClsFicProgID As String                                                                            ' L'identificateur de la class application
End Type

'******************************************************************************
'***    Type Pour La Gestion Des La Class Application                                 ***
'******************************************************************************
Public Type CLSAPL
   strClsAplProgID As String                                                                          ' L'identificateur de la class application
   strClsAplIcon As String
   strClsAplTypeFichier As String
   strClsAplOuvre As String
   strClsAplImprime As String
   strClsAplImprimeVers As String
End Type

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Definir Si Les Fichiers Et Les Applications                   ***
'***    On Ete enumerer                                                                               ***
'******************************************************************************
Private mblnClsFicAplEnumerer As Boolean                                                   ' Les Fichier Et Les References On Ete enumerer

'******************************************************************************
'***    Variable Pour Definir La Cle De Enregistrement Courant                     ***
'******************************************************************************
Private mstrClsFicCle As String                                                                    ' La Cle De Enregistrement Courant
                                                                                                                      
'******************************************************************************
'***    Variable Pour Definir L'Index De L'Enregistrement Courant                ***
'******************************************************************************
Private mlngClsFicIndex As Long                                                                    ' L'Index De L'Enregistrement Courant

'******************************************************************************
'***    Variable Pour Definir La Cle De Enregistrement Courant                    ***
'******************************************************************************
Private mstrClsAplCle As String                                                                    ' La Cle De Enregistrement Courant

'******************************************************************************
'***    Variable Pour Definir L'Index De L'Enregistrement Courant                ***
'******************************************************************************
Private mlngClsAplIndex As Long                                                                    ' L'Index De L'Enregistrement Courant

'******************************************************************************
'***    Declaration De Collection Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    Collection D'Extension De Fichier                                                      ***
'******************************************************************************
Private mcolClsFic As Collection

'******************************************************************************
'***    Collection De Reference Application                                                  ***
'******************************************************************************
Private mcolClsApl As Collection

'******************************************************************************
'***    Object Pour La  Gestion Du Registre Windows                                    ***
'******************************************************************************
Public mobjWinReg As clsRegWin                                                     ' L'object pour la gestion du registre windows
Attribute mobjWinReg.VB_VarHelpID = -1

'******************************************************************************
'***    Declaration De Evenement Public                                                      ***
'******************************************************************************

'******************************************************************************
'***    Evenement De                                                                                  ***
'******************************************************************************
Public Event FicEnumerer(ByRef ucfiClsFic As CLSFIC)

'******************************************************************************
'***    Evenement De                                                                                  ***
'******************************************************************************
Public Event AplEnumerer(ByRef ucapClsApl As CLSAPL)

'******************************************************************************
'***    Declaration De Fonction Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     FicIndex(ByVal strClsFicExtension As String) As Long                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour recuperer l'index de l'enregistrement avec la cle                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsFicExtension             - L'extension de fichier                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =FicIndex(strClsFicExtension)                                                          ***
'******************************************************************************
Private Function FicIndex(ByVal strClsFicExtension As String) As Long
   Dim blnFicOperation As Boolean                                                                  ' Le statut de l'operation
   Dim ucfiClsFic As CLSFIC                                                                            ' Un enregistrement de la collection de la class fichier
   Dim lngClsFicIndex  As Long                                                                        ' L'index de la collection de cle
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo FicIndex_Erreur
   ' Je fait defiler les index
   For lngClsFicIndex = mcolClsFic.Count To 1 Step -1
      ' Je recupere l'enregistrement
      ucfiClsFic = mcolClsFic(lngClsFicIndex)
      ' Je regarde si c'est le bon
      blnFicOperation = (ucfiClsFic.strClsFicExtension = strClsFicExtension)
      ' Je regarde si je l'ai trouver
      Select Case blnFicOperation
         Case True
            ' Je l'ai trouver
            ' Je renvoie l'index
            FicIndex = lngClsFicIndex
            ' J'arrete la boucle
            Exit For
         Case False
            ' Je ne l'ia pas trouver
      End Select
   ' Je sort de la boucle
   Next lngClsFicIndex
   ' Fin
FicIndex_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
FicIndex_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "FicIndex", _
      "", Err
   ' Je renvoie une chaine vide
   FicIndex = vbEmpty
   ' Je Continue
   Resume FicIndex_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     AplIndex(ByVal strClsAplProgID As String) As Long                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour recuperer l'index de l'enregistrement avec la cle                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsAplProgID         - L'extension de fichier                                  ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =AplIndex(strClsAplProgID)                                                            ***
'******************************************************************************
Private Function AplIndex(ByVal strClsAplProgID As String) As Long
   Dim blnAplOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucapClsApl As CLSAPL                                                                        ' Un enregistrement de la collection de class application
   Dim lngClsAplIndex  As Long                                                                       ' L'index de la collection de cle
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AplIndex_Erreur
   ' Je fait defiler les index
   For lngClsAplIndex = mcolClsApl.Count To 1 Step -1
      ' Je recupere l'enregistrement
      ucapClsApl = mcolClsApl(lngClsAplIndex)
      ' Je regarde si c'est le bon
      blnAplOperation = (ucapClsApl.strClsAplProgID = strClsAplProgID)
      ' Je regarde si je l'ai trouver
      Select Case blnAplOperation
         Case True
            ' Je l'ai trouver
            ' Je renvoie l'index
            AplIndex = lngClsAplIndex
            ' J'arrete la boucle
            Exit For
         Case False
            ' Je ne l'ia pas trouver
      End Select
   ' Je sort de la boucle
   Next lngClsAplIndex
   ' Fin
AplIndex_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
AplIndex_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "AplIndex", _
      "", Err
   ' Je renvoie une chaine vide
   AplIndex = vbEmpty
   ' Je Continue
   Resume AplIndex_Exit
   ' Fin
End Function

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     LectureFic(ByVal strClsFicExtension As String,                                ***
'***     ByRef strClsFicProgID As String) As Boolean                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'**       Pour lire la reference de l'extension de fichier                                  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsFicExtension      - L'extension de fichier                                 ***
'***      strClsFicProgID           - La reference de l'extension de fichier          ***
'***    SORTIE:                                                                                           ***
'***      strClsFicProgID          - La reference de l'extension de fichier          ***
'***      LectureFic                  - Retourne la reference                                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =LectureFic(".XXX",strClsFicProgID)                                                ***
'******************************************************************************
Public Function LectureFic(ByVal strClsFicExtension As String, ByRef strClsFicProgID As String) As Boolean
   Dim blnFicOperation As Boolean                                                                 ' Le statut de l'operation
   Dim strrcsClassFichier As String                                                                ' La sous cle de la base de registre
   Dim ucfiClsFic As CLSFIC                                                                           ' Un enregistrement de la collection de class fichier
   Dim vntucfiClsFic As Variant                                                                      ' Un enregistrement de la collection de class fichier
   Dim vntClsFicProgID As Variant                                                                  ' La valeur de la reference
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucfiClsFic = mcolClsFic(strClsFicExtension)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureFic_Erreur
   ' Je regarde la valeur
   blnFicOperation = Not IsEmpty(vntucfiClsFic)
   ' Je regarde si j'ai un enregistrement
   Select Case blnFicOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucfiClsFic = vntucfiClsFic
         ' Je regarde si c'est l'enregistrement courant
         blnFicOperation = (mstrClsFicCle = ucfiClsFic.strClsFicExtension)
         ' Je regarde la valeur
         Select Case blnFicOperation
            Case True
               ' C'est  l'enregistrement courant
            Case False
               ' Ce n'est pas l'enregistrement courant
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsFicIndex = FicIndex(ucfiClsFic.strClsFicExtension)
               ' Je met a jour la cle de la cle de l'enregistrement courant de la collection
               mstrClsFicCle = ucfiClsFic.strClsFicExtension
         End Select
         ' Je renvoie la valeur
         LectureFic = True
      Case False
         ' C'est une valeur vide
         ' Je defini l'enregistrement
         ucfiClsFic.strClsFicExtension = strClsFicExtension
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsFicIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsFicCle = ucfiClsFic.strClsFicExtension
         ' Je regarde la valeur de la cle racine des extension
         Select Case mconrcsClassFichier
            Case vbNullString
              ' Il n'y a pas de sous cle
              ' Je met la valeur a jour
              strrcsClassFichier = vbNullString
            Case Else
              ' Il y a une sous cle
              ' Je met la valeur a jour
              strrcsClassFichier = mconrcsClassFichier
         End Select
         ' Je recupere  la valeur
         LectureFic = mobjWinReg.LectureValeur(mconrcrClassFichier, strrcsClassFichier & "\" & ucfiClsFic.strClsFicExtension, "", vntClsFicProgID)
         ' Je regarde si la valeur n'est pas null
         Select Case LectureFic
            Case True
               ' La valeur n'est pas vide
               ' Je recupere la reference
               ucfiClsFic.strClsFicProgID = vntClsFicProgID
               ' Je l'ajoute dans la collection des extension
               mcolClsFic.Add ucfiClsFic, ucfiClsFic.strClsFicExtension
               ' Je recupere l'index
               mlngClsFicIndex = mcolClsFic.Count
            Case False
               ' La valeur est vide
               ' Je defini l'enregistrement
               ucfiClsFic.strClsFicExtension = vbNullString
               ' Je met une chaine vide pour la reference
               ucfiClsFic.strClsFicProgID = vbNullString
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsFicIndex = vbEmpty
               ' Je met a jour la cle de l'enregistrement courant de la collection
               mstrClsFicCle = ucfiClsFic.strClsFicExtension
         End Select
   End Select
   ' Je defini la reference
   strClsFicProgID = ucfiClsFic.strClsFicProgID
   ' Fin
LectureFic_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
LectureFic_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "LectureFic", strClsFicExtension & "," & strClsFicProgID, Err
   ' Je renvoie que j'ai bien lu
   LectureFic = False
   ' Je Continue
   Resume LectureFic_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EcritureFic(ByVal strClsFicExtension As String,                               ***
'***     ByVal strClsFicProgID As String) As Boolean                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'**       Pour ecrire l'extension et la reference de fichier                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsFicExtension     - L'extension de fichier                                  ***
'***      strClsFicProgID          - La reference de l'extension de fichier           ***
'***    SORTIE:                                                                                           ***
'***      EcritureFic                 - Retourne le resultat  de l'ecriture reference ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EcritureFic(".XXX","Fichier XXX")                                                 ***
'******************************************************************************
Public Function EcritureFic(ByVal strClsFicExtension As String, ByVal strClsFicProgID As String) As Boolean
Attribute EcritureFic.VB_Description = "Pour ecrire une extension"
   Dim blnFicOperation As Boolean                                                                 ' Le statut de l'operation
   Dim strrcsClassFichier As String                                                                ' La sous cle de la base de registre
   Dim ucfiClsFic As CLSFIC                                                                           ' Un enregistrement de la collection d'extension de fichier
   Dim vntucfiClsFic As Variant                                                                      ' Un enregistrement de la collection d'extension de fichier
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucfiClsFic = mcolClsFic(strClsFicProgID)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureFic_Erreur
   ' Je regarde la valeur
   blnFicOperation = Not IsEmpty(vntucfiClsFic)
   ' Je regarde si j'ai un enregistrement
   Select Case blnFicOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucfiClsFic = vntucfiClsFic
         ' Je regarde si c'est l'enregistrement courant
         blnFicOperation = (mstrClsFicCle = ucfiClsFic.strClsFicProgID)
         ' Je regarde la valeur
         Select Case blnFicOperation
            Case True
               ' C'est  l'enregistrement courant
            Case False
               ' Ce n'est pas l'enregistrement courant
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsFicIndex = FicIndex(ucfiClsFic.strClsFicProgID)
               ' Je met a jour la cle de la cle de l'enregistrement courant de la collection
               mstrClsFicCle = ucfiClsFic.strClsFicProgID
         End Select
         ' Je l'ajoute dans la collection des extension
         mcolClsFic.Remove ucfiClsFic.strClsFicProgID
      Case False
         ' C'est une valeur vide
         ' Je defini l'enregistrement
         ucfiClsFic.strClsFicProgID = strClsFicProgID
         ' Je defini l'enregistrement
         ucfiClsFic.strClsFicProgID = strClsFicProgID
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsFicIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsFicCle = ucfiClsFic.strClsFicProgID
   End Select
    ' Je regarde la valeur de la cle racine des extension
   Select Case mconrcsClassFichier
      Case vbNullString
        ' Il n'y a pas de sous cle
        ' Je met la valeur a jour
        strrcsClassFichier = ucfiClsFic.strClsFicProgID
      Case Else
        ' Il y a une sous cle
        ' Je met la valeur a jour
        strrcsClassFichier = mconrcsClassFichier & "\" & ucfiClsFic.strClsFicProgID
   End Select
   ' J'ecrit l'extension
   mobjWinReg.EcritureCle mconrcrClassFichier, strrcsClassFichier
   ' J'ecrit la reference
   EcritureFic = mobjWinReg.EcritureValeur(mconrcrClassFichier, strrcsClassFichier, "", ucfiClsFic.strClsFicProgID)
   ' Je regarde si la valeur n'est pas null
   Select Case EcritureFic
      Case True
         ' La valeur n'est pas vide
         ' Je regarde l'index de cle courant
         Select Case mlngClsFicIndex
            Case vbEmpty
               ' L'index de cle courant est a zero
               ' Je l'ajoute dans la collection des extension
               mcolClsFic.Add ucfiClsFic, ucfiClsFic.strClsFicProgID
               ' Je recupere l'index
               mlngClsFicIndex = mcolClsFic.Count
            Case Else
               ' L'index de cle courant n'est pas a zero
               ' Je la rajoute dans la collection
               mcolClsFic.Add ucfiClsFic, ucfiClsFic.strClsFicProgID, , mlngClsFicIndex - 1
         End Select
      Case False
         ' La valeur est vide
         ' Je defini l'enregistrement
         ucfiClsFic.strClsFicProgID = vbNullString
         ' Je met une chaine vide pour la reference
         ucfiClsFic.strClsFicProgID = vbNullString
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsFicIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsFicCle = ucfiClsFic.strClsFicProgID
   End Select
   ' Fin
EcritureFic_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EcritureFic_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EcritureFic", strClsFicProgID & "," & strClsFicProgID, Err
   ' Je renvoie que faux
   EcritureFic = False
   ' Je Continue
   Resume EcritureFic_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     AjouteExt(ByVal strClsFicProgID As String,                                         ***
'***     ByVal strClsFicProgID As String) As Boolean                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ajouter une extension de fichier                                                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsFicProgID           - La valeur de l'extension                                 ***
'***      strClsFicExtension   - La reference de l'extension                            ***
'***    SORTIE:                                                                                            ***
'***      AjouteExt                 - Retourne si l'extension a ete ajouter              ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =AjouteExt(".XXX","Fichier XXX")                                                   ***
'******************************************************************************
Public Function AjouteExt(ByVal strClsFicProgID As String, ByVal strClsFicExtension As String) As Boolean
   Dim ucfiClsFic As CLSFIC                                                                           ' Un enregistrement de la collection des extension de fichier
   Dim strrcsClassFichier As String                                                                     ' La sous cle de la base de registre
  ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AjouteExt_Erreur
   ' Je defini l'enregistrement
   ucfiClsFic.strClsFicProgID = strClsFicProgID
   ' Je defini l'enregistrement
   ucfiClsFic.strClsFicExtension = strClsFicExtension
   ' Je regarde la valeur de la cle racine des extension
   Select Case mconrcsClassFichier
      Case vbNullString
        ' Il n'y a pas de sous cle
        ' Je met la valeur a jour
        strrcsClassFichier = ucfiClsFic.strClsFicProgID
      Case Else
        ' Il y a une sous cle
        ' Je met la valeur a jour
        strrcsClassFichier = mconrcsClassFichier & "\" & ucfiClsFic.strClsFicProgID
   End Select
   ' Je regarde si l'extension existe
   AjouteExt = mobjWinReg.OuvreCle(mconrcrClassFichier, strrcsClassFichier)
   ' Je regarde le resultat
   Select Case AjouteExt
      Case True
         ' L'extension existe
         ' Je renvoie la valeur
         AjouteExt = False
      Case False
         ' L'extension existe
         ' Je l'ecrit
         AjouteExt = EcritureFic(ucfiClsFic.strClsFicProgID, ucfiClsFic.strClsFicExtension)
   End Select
   ' Fin
AjouteExt_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
AjouteExt_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "AjouteExt", strClsFicProgID & "," & strClsFicProgID, Err
   ' Je renvoie une chaine vide
   AjouteExt = False
   ' Je Continue
   Resume AjouteExt_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     SupprimeExt(ByVal strClsFicProgID As String) As Boolean                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer une extension de fichier                                           ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsFicProgID       - La valeur de l'extension                                    ***
'***    SORTIE:                                                                                           ***
'***      SupprimeExt        - Le resultat de la supression                               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =SupprimeExt(".XXX")                                                                     ***
'******************************************************************************
Public Function SupprimeExt(ByVal strClsFicProgID As String) As Boolean
   Dim blnFicOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucfiClsFic As CLSFIC                                                                           ' Un enregistrement de la collection d'extension de fichier
   Dim vntucfiClsFic As Variant                                                                      ' Un enregistrement de la collection d'extension de fichier
   Dim strrcsClassFichier As String                                                                     ' La sous cle de la base de registre
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucfiClsFic = mcolClsFic(strClsFicProgID)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo SupprimeExt_Erreur
   ' Je regarde la valeur
   blnFicOperation = Not IsEmpty(vntucfiClsFic)
   ' Je regarde si j'ai un enregistrement
   Select Case blnFicOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucfiClsFic = vntucfiClsFic
         ' Je regarde la valeur de la cle racine des extension
         Select Case mconrcsClassFichier
            Case vbNullString
               ' Il n'y a pas de sous cle
               ' Je met la valeur a jour
               strrcsClassFichier = ucfiClsFic.strClsFicProgID
            Case Else
               ' Il y a une sous cle
               ' Je met la valeur a jour
               strrcsClassFichier = mconrcsClassFichier & "\" & ucfiClsFic.strClsFicProgID
         End Select
         ' J'ecrit l'extension
         SupprimeExt = mobjWinReg.SupprimeCle(mconrcrClassFichier, strrcsClassFichier)
         ' Je regarde si la valeur n'est pas null
         Select Case SupprimeExt
            Case True
               ' La valeur n'est pas vide
               ' Je l'enleve dans la collection des extension
               mcolClsFic.Remove ucfiClsFic.strClsFicProgID
            Case False
               ' La valeur est vide
         End Select
         ' Je defini l'enregistrement
         ucfiClsFic.strClsFicProgID = vbNullString
         ' Je met une chaine vide pour la reference
         ucfiClsFic.strClsFicProgID = vbNullString
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsFicIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsFicCle = ucfiClsFic.strClsFicProgID
      Case False
         ' C'est une valeur vide
         ' Je renvoie la valeur
         SupprimeExt = True
   End Select
   ' Fin
SupprimeExt_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
SupprimeExt_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "SupprimeExt", strClsFicProgID, Err
   ' Je renvoie que faux
   SupprimeExt = False
   ' Je Continue
   Resume SupprimeExt_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     LectureApl(ByVal strClsAplProgID As String,                                   ***
'***     ByRef strClsAplOuvre As String,                                                     ***
'***     ByRef strClsAplIcon As String) As Boolean                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'**       Pour lire une reference d'application et ces command                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsAplProgID                        - Une reference d'application          ***
'***      strClsAplOuvre                         - La command ouvrir                       ***
'***      strClsAplIcon                            - La command icon                         ***
'***    SORTIE:                                                                                           ***
'***      strClsAplOuvre                         - La command ouvrir                      ***
'***      strClsAplIcon                        - La command icon                            ***
'***      LectureApl                             - Retourne le resultat                         ***
'***                                                     de la lecture reference                     ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =LectureApl("Fichier XXX",strClsAplOuvre,                                     ***
'***     strClsAplIcon)                                                                                ***
'******************************************************************************
Public Function LectureApl(ByVal strClsAplProgID As String, ByRef strClsAplOuvre As String, ByRef strClsAplIcon As String) As Boolean
   Dim blnAplOperation As Boolean                                                                 ' Le statut de l'operation
   Dim strrcsApplication As String                                                                    ' La sous cle de la base de registre
   Dim ucapClsApl As CLSAPL                                                                        ' Un enregistrement de la collection de reference application
   Dim vntucapClsApl As Variant                                                                    ' Un enregistrement de la collection de reference application
   Dim vntClsAplCommand As Variant                                                            ' La valeur de la commande
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucapClsApl = mcolClsApl(strClsAplProgID)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureApl_Erreur
   ' Je regarde la valeur
   blnAplOperation = Not IsEmpty(vntucapClsApl)
   ' Je regarde si j'ai un enregistrement
   Select Case blnAplOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucapClsApl = vntucapClsApl
         ' Je regarde si c'est l'enregistrement courant
         blnAplOperation = (mstrClsAplCle = ucapClsApl.strClsAplProgID)
         ' Je regarde la valeur
         Select Case blnAplOperation
            Case True
               ' C'est  l'enregistrement courant
            Case False
               ' Ce n'est pas l'enregistrement courant
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsAplIndex = AplIndex(ucapClsApl.strClsAplProgID)
               ' Je met a jour la cle de la cle de l'enregistrement courant de la collection
               mstrClsAplCle = ucapClsApl.strClsAplProgID
         End Select
         ' Je renvoie la valeur
         LectureApl = True
      Case False
         ' C'est une valeur vide
         ' Je defini l'enregistrement
         ucapClsApl.strClsAplProgID = strClsAplProgID
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsAplIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsAplCle = ucapClsApl.strClsAplProgID
         ' Je regarde la valeur de la cle racine des extension
         Select Case mconrcsClassApplication
            Case vbNullString
              ' Il n'y a pas de sous cle
              ' Je met la valeur a jour
              strrcsApplication = vbNullString
            Case Else
              ' Il y a une sous cle
              ' Je met la valeur a jour
              strrcsApplication = mconrcsClassApplication
         End Select
         ' Je lit la valeur de la commande
         blnAplOperation = mobjWinReg.LectureValeur(mconrcrClassApplication, _
            strrcsApplication & "\" & ucapClsApl.strClsAplProgID & "\" & mconrcsIcon, "", vntClsAplCommand)
         ' Je regarde si je l'ai recuperer
         Select Case blnAplOperation
            Case True
               ' La valeur n'est pas vide
               ' Je ferme la cle
               mobjWinReg.FermeCle mconrcrClassApplication, strrcsApplication & "\" & ucapClsApl.strClsAplProgID & "\" & mconrcsIcon
               ' Je recupere la reference
               ucapClsApl.strClsAplIcon = vntClsAplCommand
            Case False
               ' La valeur est vide
               ' Je met une chaine vide pour la reference
               ucapClsApl.strClsAplIcon = vbNullString
         End Select
         ' Je lit la valeur de la commande
         LectureApl = mobjWinReg.LectureValeur(mconrcrClassApplication, _
            strrcsApplication & "\" & ucapClsApl.strClsAplProgID & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand, "", vntClsAplCommand)
         ' Je regarde si je l'ai recuperer
         Select Case LectureApl
            Case True
               ' La valeur n'est pas vide
               ' Je ferme la cle
               mobjWinReg.FermeCle mconrcrClassApplication, strrcsApplication & "\" & ucapClsApl.strClsAplProgID & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand
               ' Je recupere la reference
               ucapClsApl.strClsAplOuvre = vntClsAplCommand
               ' Je regarde l'index de cle courant
               Select Case mlngClsAplIndex
                  Case vbEmpty
                     ' L'index de cle courant est a zero
                     ' Je l'ajoute dans la collection des reference
                     mcolClsApl.Add ucapClsApl, ucapClsApl.strClsAplProgID
                     ' Je recupere l'index
                     mlngClsAplIndex = mcolClsApl.Count
                  Case Else
                     ' L'index de cle courant n'est pas a zero
                     ' Je la rajoute dans la collection
                     mcolClsApl.Add ucapClsApl, ucapClsApl.strClsAplProgID, mlngClsAplIndex
               End Select
            Case False
               ' La valeur est vide
               ' Je met une chaine vide pour la reference
               ucapClsApl.strClsAplProgID = vbNullString
               ' Je met une chaine vide pour la reference
               ucapClsApl.strClsAplOuvre = vbNullString
               ' Je met une chaine vide pour la reference
               ucapClsApl.strClsAplIcon = vbNullString
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsAplIndex = vbEmpty
               ' Je met a jour la cle de l'enregistrement courant de la collection
               mstrClsAplCle = ucapClsApl.strClsAplProgID
         End Select
   End Select
   ' Je defini la reference
   strClsAplOuvre = ucapClsApl.strClsAplOuvre
   ' Je defini la reference
   strClsAplIcon = ucapClsApl.strClsAplIcon
   ' Fin
LectureApl_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
LectureApl_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "LectureApl", strClsAplProgID & "," & strClsAplOuvre + "," + strClsAplIcon, Err
   ' Je renvoie que faux
   LectureApl = False
   ' Je Continue
   Resume LectureApl_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EcritureRef(ByVal strClsAplProgID As String,                                     ***
'***     ByVal strCLSAPLCommandOuvre As String,                                      ***
'***     ByVal strCLSAPLCommandIcon As String) As Boolean                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'**       Pour ecrire une reference d'application et ces command                    ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsAplProgID                        - Une reference d'application             ***
'***      strCLSAPLCommandOuvre       - La command ouvrir                         ***
'***      strCLSAPLCommandIcon         - La command icon                            ***
'***    SORTIE:                                                                                           ***
'***      EcritureRef                           - Retourne le resultat                         ***
'***                                                     de la ecriture reference                    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EcritureRef("Fichier XXX","C:\COMMAND.COM",                            ***
'***     "C:\COMMAND.COM")                                                                      ***
'******************************************************************************
Public Function EcritureRef(ByVal strClsAplProgID As String, ByVal strCLSAPLCommandOuvre As String, ByVal strCLSAPLCommandIcon As String) As Boolean
Attribute EcritureRef.VB_Description = "Pour ecrire une reference"
   Dim blnAplOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucapClsApl As CLSAPL                                                                        ' Un enregistrement de la collection de reference application
   Dim vntucapClsApl As Variant                                                                     ' Un enregistrement de la collection de reference application
   Dim strrcsReference As String                                                                     ' La sous cle de la base de registre
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucapClsApl = mcolClsApl(strClsAplProgID)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureRef_Erreur
   ' Je regarde la valeur
   blnAplOperation = Not IsEmpty(vntucapClsApl)
   ' Je regarde si j'ai un enregistrement
   Select Case blnAplOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucapClsApl = vntucapClsApl
         ' Je regarde si c'est l'enregistrement courant
         blnAplOperation = (mstrClsAplCle = ucapClsApl.strClsAplProgID)
         ' Je regarde la valeur
         Select Case blnAplOperation
            Case True
               ' C'est  l'enregistrement courant
            Case False
               ' Ce n'est pas l'enregistrement courant
               ' Je met a jour l'index de cle de l'enregistrement courant de la collection
               mlngClsAplIndex = AplIndex(ucapClsApl.strClsAplProgID)
               ' Je met a jour la cle de la cle de l'enregistrement courant de la collection
               mstrClsAplCle = ucapClsApl.strClsAplProgID
         End Select
         ' Je l'ajoute dans la collection des references
         mcolClsApl.Remove ucapClsApl.strClsAplProgID
      Case False
         ' C'est une valeur vide
         ' Je defini l'enregistrement
         ucapClsApl.strClsAplProgID = strClsAplProgID
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsAplIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsAplCle = ucapClsApl.strClsAplProgID
   End Select
   ' Je defini l'enregistrement
   ucapClsApl.strClsAplOuvre = strCLSAPLCommandOuvre
   ' Je defini l'enregistrement
   ucapClsApl.strClsAplIcon = strCLSAPLCommandIcon
   ' Je regarde la valeur de la cle racine des extension
   Select Case mconrcsClassApplication
      Case vbNullString
        ' Il n'y a pas de sous cle
        ' Je met la valeur a jour
        strrcsReference = ucapClsApl.strClsAplProgID
      Case Else
        ' Il y a une sous cle
        ' Je met la valeur a jour
        strrcsReference = mconrcrClassApplication & "\" & ucapClsApl.strClsAplProgID
   End Select
   ' J'ecrit la cle de la reference application
   mobjWinReg.EcritureCle mconrcrClassApplication, strrcsReference
   ' J'ecrit la cle de command icon
   mobjWinReg.EcritureCle mconrcrClassApplication, strrcsReference & "\" & mconrcsIcon
   ' J'ecrit la valeur command icon
   mobjWinReg.EcritureValeur mconrcrClassApplication, strrcsReference & "\" & mconrcsIcon, "", ucapClsApl.strClsAplIcon
   ' Je ferme la cle de command icon
   mobjWinReg.FermeCle mconrcrClassApplication, strrcsReference & "\" & mconrcsIcon
   ' J'ecrit la cle de command ouvrir
   mobjWinReg.EcritureCle mconrcrClassApplication, _
      strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand
   ' J'ecrit la valeur command ouvrir
   EcritureRef = mobjWinReg.EcritureValeur(mconrcrClassApplication, _
      strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand, "", ucapClsApl.strClsAplOuvre)
   ' Je ferme la cle de command ouvrir
   mobjWinReg.FermeCle mconrcrClassApplication, _
      strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand
   ' Je regarde si je l'ai recuperer
   Select Case EcritureRef
      Case True
         ' La valeur n'est pas vide
         ' Je regarde l'index de cle courant
         Select Case mlngClsAplIndex
            Case vbEmpty
               ' L'index de cle courant est a zero
               ' Je l'ajoute dans la collection des reference
               mcolClsApl.Add ucapClsApl, ucapClsApl.strClsAplProgID
               ' Je recupere l'index
               mlngClsAplIndex = mcolClsApl.Count
            Case Else
               ' L'index de cle courant n'est pas a zero
               ' Je la rajoute dans la collection
               mcolClsApl.Add ucapClsApl, ucapClsApl.strClsAplProgID, , mlngClsAplIndex - 1
         End Select
      Case False
         ' La valeur est vide
         ' Je met une chaine vide pour la reference
         ucapClsApl.strClsAplProgID = vbNullString
         ' Je met une chaine vide pour la reference
         ucapClsApl.strClsAplOuvre = vbNullString
         ' Je met une chaine vide pour la reference
         ucapClsApl.strClsAplIcon = vbNullString
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsAplIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsAplCle = ucapClsApl.strClsAplProgID
   End Select
   ' Fin
EcritureRef_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EcritureRef_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EcritureRef", strClsAplProgID & "," & strCLSAPLCommandOuvre & "," & strCLSAPLCommandIcon, Err
   ' Je renvoie que faux
   EcritureRef = False
   ' Je Continue
   Resume EcritureRef_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     AjouteRef(ByVal strClsAplProgID As String,                                        ***
'***     ByVal strCLSAPLCommandOuvre As String,                                       ***
'***     ByVal strCLSAPLCommandIcon As String) As Boolean                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ajouter une reference application                                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsAplProgID                     - La nom  de la reference                      ***
'***      strCLSAPLCommandOuvre    - La command ouvrir de la reference     ***
'***      strCLSAPLCommandIcon    - La command ouvrir de la reference        ***
'***    SORTIE:                                                                                           ***
'***      AjouteRef                            - Retourne si la reference a ete ajouter ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     Me.AjouteRef("Fichier XXX","C:\COMMAND.COM %1",                      ***
'***     "C:\COMMAND.COM,1")                                                                    ***
'******************************************************************************
Public Function AjouteRef(ByVal strClsAplProgID As String, ByVal strCLSAPLCommandOuvre As String, ByVal strCLSAPLCommandIcon As String) As Boolean
   Dim urappCLSAPL As CLSAPL                                                                       ' Un enregistrement de la collection
   Dim strrcsReference As String                                                                     ' La sous cle de la base de registre
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AjouteRef_Erreur
   ' Je defini l'enregistrement
   urappCLSAPL.strClsAplProgID = strClsAplProgID
   ' Je defini l'enregistrement
   urappCLSAPL.strClsAplOuvre = strCLSAPLCommandOuvre
   ' Je defini l'enregistrement
   urappCLSAPL.strClsAplIcon = strCLSAPLCommandIcon
   ' Je regarde la valeur de la cle racine des extension
   Select Case mconrcsClassApplication
      Case vbNullString
        ' Il n'y a pas de sous cle
        ' Je met la valeur a jour
        strrcsReference = urappCLSAPL.strClsAplProgID
      Case Else
        ' Il y a une sous cle
        ' Je met la valeur a jour
        strrcsReference = mconrcsClassApplication & "\" & urappCLSAPL.strClsAplProgID
   End Select
   ' Je regarde si l'extension existe
   AjouteRef = mobjWinReg.OuvreCle(mconrcrClassApplication, strrcsReference)
   ' Je regarde le resultat
   Select Case AjouteRef
      Case True
         ' L'extension existe
         ' Je renvoie la valeur
         AjouteRef = False
      Case False
         ' L'extension n'existe pas
         ' Je l'ecrit
         AjouteRef = EcritureRef(urappCLSAPL.strClsAplProgID, urappCLSAPL.strClsAplOuvre, urappCLSAPL.strClsAplIcon)
   End Select
   ' Fin
AjouteRef_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
AjouteRef_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "AjouteRef", strClsAplProgID & "," & strCLSAPLCommandOuvre + "," + strCLSAPLCommandIcon, Err
   ' Je renvoie une chaine vide
   AjouteRef = False
   ' Je Continue
   Resume AjouteRef_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     SupprimeRef(ByVal strClsFicProgID As String) As Boolean                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer une reference application                                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strClsAplProgID       - La valeur de la reference                                 ***
'***    SORTIE:                                                                                           ***
'***      SupprimeRef        - Le resultat de la supression                               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =SupprimeRef("Fichier XXX")                                                           ***
'******************************************************************************
Public Function SupprimeRef(ByVal strClsAplProgID As String) As Boolean
   Dim blnAplOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucapClsApl As CLSAPL                                                                       ' Un enregistrement de la collection de reference application
   Dim vntucapClsApl As Variant                                                                   ' Un enregistrement de la collection de reference application
   Dim strrcsReference As String                                                                    ' La sous cle de la base de registre
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucapClsApl = mcolClsApl(strClsAplProgID)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo SupprimeRef_Erreur
   ' Je regarde la valeur
   blnAplOperation = Not IsEmpty(vntucapClsApl)
   ' Je regarde si j'ai un enregistrement
   Select Case blnAplOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je renvoie la valeur
         ucapClsApl = vntucapClsApl
         ' Je regarde la valeur de la cle racine des extension
         Select Case strrcsReference
            Case vbNullString
               ' Il n'y a pas de sous cle
               ' Je met la valeur a jour
               strrcsReference = ucapClsApl.strClsAplProgID
            Case Else
               ' Il y a une sous cle
               ' Je met la valeur a jour
               strrcsReference = strrcsReference & "\" & ucapClsApl.strClsAplProgID
         End Select
         ' J'ouvre la cle command
         mobjWinReg.OuvreCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand
         ' Je supprime la cle command
         mobjWinReg.SupprimeCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre & "\" & mconrcsCommand
         ' J'ouvre la cle ouvre
         mobjWinReg.OuvreCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre
         ' Je supprime la cle ouvre
         mobjWinReg.SupprimeCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau & "\" & mconrcsOuvre
         ' J'ouvre la cle shell
         mobjWinReg.OuvreCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau
         ' Je supprime la cle shell
         mobjWinReg.SupprimeCle mconrcrClassApplication, _
            strrcsReference & "\" & mconrcsBureau
         ' J'ouvre la cle icon
         mobjWinReg.OuvreCle mconrcrClassApplication, strrcsReference & "\" & mconrcsIcon
         ' Je supprime la cle icon
         mobjWinReg.SupprimeCle mconrcrClassApplication, strrcsReference & "\" & mconrcsIcon
         ' Je supprime la reference
         SupprimeRef = mobjWinReg.SupprimeCle(mconrcrClassApplication, strrcsReference)
         ' Je regarde si la valeur n'est pas null
         Select Case SupprimeRef
            Case True
               ' La valeur n'est pas vide
               ' Je l'enleve dans la collection des extension
               mcolClsApl.Remove ucapClsApl.strClsAplProgID
            Case False
               ' La valeur est vide
         End Select
         ' Je defini l'enregistrement
         ucapClsApl.strClsAplProgID = vbNullString
         ' Je met une chaine vide pour la reference
         ucapClsApl.strClsAplOuvre = vbNullString
         ' Je met une chaine vide pour la reference
         ucapClsApl.strClsAplIcon = vbNullString
         ' Je met a jour l'index de cle de l'enregistrement courant de la collection
         mlngClsAplIndex = vbEmpty
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsAplCle = ucapClsApl.strClsAplProgID
      Case False
         ' C'est une valeur vide
         ' Je renvoie la valeur
         SupprimeRef = True
   End Select
   ' Fin
SupprimeRef_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
SupprimeRef_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "SupprimeRef", strClsAplProgID, Err
   ' Je renvoie que faux
   SupprimeRef = False
   ' Je Continue
   Resume SupprimeRef_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Enumere() As Boolean                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour enumerer une collection d'extension de fichier                          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Enumere   - Retourne le resultat de l'operation                                ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =Me.Enumere                                                                                  ***
'******************************************************************************
Public Function Enumere() As Boolean
Attribute Enumere.VB_Description = "Pour enumerer les extension et les references des fichiers"
  ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Enumere_Erreur
   ' Je regarde si j'ai deja enumerer les cles
   Select Case mblnClsFicAplEnumerer
      Case True
         ' J'ai deja enumerer le cles
      Case False
         ' Je n'ai pas enumerer le cles
         ' J'enumere les cles extensions
         mblnClsFicAplEnumerer = mobjWinReg.EnumereCle(mconrcrClassFichier, mconrcsClassFichier)
   End Select
   ' Je retourne le resultat
   Enumere = mblnClsFicAplEnumerer
   ' Fin
Enumere_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
Enumere_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "Enumere", "", Err
   ' Je renvoie une chaine vide
   Enumere = False
   ' Je Continue
   Resume Enumere_Exit
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
'***     Pour creer une collection d'extension a la creation                           ***
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
   ' J'initialise la collection d'extension de fichier
   Set mcolClsFic = New Collection
   ' J'initialise la collection de reference application
   Set mcolClsApl = New Collection
   ' Jinitialise l'object registre
   Set mobjWinReg = New clsRegWin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'***     Class_Terminate                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer la collection d'extension a la suppression                 ***
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
   ' J'initialise la collection d'extension de fichier
   Set mcolClsFic = Nothing
   ' J'initialise la collection de reference application
   Set mcolClsApl = Nothing
   ' Jinitialise l'object registre
   Set mobjWinReg = Nothing
End Sub

'******************************************************************************
'***    Declaration De Propriete Public                                                         ***
'******************************************************************************

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***     FicCompte() As Long                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour renvoyer le nombre d'enregistrement de la collection                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     Neant                                                                                              ***
'***    SORTIE:                                                                                           ***
'***     FicCompte                        - La valeur du nombre d'enregistrement       ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =Me.FicCompte(1)                                                                               ***
'******************************************************************************
Public Property Get FicCompte() As Long
Attribute FicCompte.VB_Description = "Le nombre d'extension enumerer"
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo FicCompte_Erreur
   ' Je renvoie le nombre d'enregistrement
    FicCompte = mcolClsFic.Count
   ' Fin
FicCompte_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
FicCompte_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROPRIETE, "FicCompte", "", Err
   ' Je renvoie zero a l'user
   FicCompte = vbEmpty
   ' Je Continue
   Resume FicCompte_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***     Fic(ByVal lngClsFicIndex As Long) As CLSFIC                                  ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour renvoyer l'extension de la collection                                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     lngClsFicIndex          - La valeur de l'extension                                 ***
'***    SORTIE:                                                                                           ***
'***     Ext                            - La valeur de l'extension                                ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =Me.Fic(1)                                                                                       ***
'******************************************************************************
Public Property Get Fic(ByVal lngClsFicIndex As Long) As CLSFIC
Attribute Fic.VB_Description = "La valeur de l'extension"
   Dim blnFicOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucfiClsFic As CLSFIC                                                                           ' Un enregistrement de la collection d'extension
   Dim vntucfiClsFic As Variant                                                                      ' Un enregistrement de la collection d'extension
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucfiClsFic = mcolClsFic(lngClsFicIndex)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Fic_Erreur
   ' Je regarde si il y a un enregistrement
   blnFicOperation = Not IsEmpty(vntucfiClsFic)
   ' Je regarde si j'ai un enregistrement
   Select Case blnFicOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je recupere l'enregistrement
         ucfiClsFic = vntucfiClsFic
         ' Je met a jour l'index de cle de l'enregistrement courant
         mlngClsFicIndex = lngClsFicIndex
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsFicCle = ucfiClsFic.strClsFicExtension
      Case False
         ' C'est une valeur vide
   End Select
   ' Je renvoie la valeur
   Fic = ucfiClsFic
   ' Fin
Fic_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
Fic_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROPRIETE, "Fic", CStr(lngClsFicIndex), Err
   ' Je renvoie une chaine vide
   Fic = ucfiClsFic
   ' Je Continue
   Resume Fic_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***     AplCompte() As Long                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour renvoyer le nombre d'enregistrement de la collection                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     Neant                                                                                              ***
'***    SORTIE:                                                                                           ***
'***     AplCompte                       - La valeur du nombre d'enregistrement       ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =Me.AplCompte(1)                                                                              ***
'******************************************************************************
Public Property Get AplCompte() As Long
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AplCompte_Erreur
   ' Je renvoie le nombre d'enregistrement
    AplCompte = mcolClsApl.Count
  ' Fin
AplCompte_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
AplCompte_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROPRIETE, "AplCompte", "", Err
   ' Je renvoie zero a l'user
   AplCompte = vbEmpty
   ' Je Continue
   Resume AplCompte_Exit
   ' Fin
End Property

'******************************************************************************
'***    PROPRIETE:                                                                                     ***
'***     Apl(ByVal lngClsAplIndex As Long) As CLSAPL                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour renvoyer la reference  de la collection                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     lngClsAplIndex          - L'index de la reference                                  ***
'***    SORTIE:                                                                                           ***
'***     Apl                            - La valeur de la reference                              ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =Me.Apl(1)                                                                                       ***
'******************************************************************************
Public Property Get Apl(ByVal lngClsAplIndex As Long) As CLSAPL
Attribute Apl.VB_Description = "La valeur de la reference"
   Dim blnAplOperation As Boolean                                                                 ' Le statut de l'operation
   Dim ucapClsApl As CLSAPL                                                                        ' Un enregistrement de la collection de reference
   Dim vntucapClsApl As Variant                                                                     ' Un enregistrement de la collection de reference
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je recupere l'enregistrement si elle  existe
   vntucapClsApl = mcolClsApl(lngClsAplIndex)
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Apl_Erreur
   ' Je regarde si il y a un enregistrement
   blnAplOperation = Not IsEmpty(vntucapClsApl)
   ' Je regarde si j'ai un enregistrement
   Select Case blnAplOperation
      Case True
         ' Ce n'est pas une valeur vide
         ' Je recupere l'enregistrement
         ucapClsApl = vntucapClsApl
         ' Je met a jour l'index de cle de l'enregistrement courant
         mlngClsAplIndex = lngClsAplIndex
         ' Je met a jour la cle de l'enregistrement courant de la collection
         mstrClsAplCle = ucapClsApl.strClsAplProgID
      Case False
         ' C'est une valeur vide
   End Select
   ' Je renvoie la valeur
   Apl = ucapClsApl
   ' Fin
Apl_Exit:
   ' Je Sort De La Procedure
   Exit Property
   ' Fin
Apl_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEPROPRIETE, "Apl", CStr(lngClsAplIndex), Err
   ' Je renvoie une chaine vide
   Apl = ucapClsApl
   ' Je Continue
   Resume Apl_Exit
   ' Fin
End Property

'******************************************************************************
'***    Evenement                                                                                       ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***     mobjWinReg_CleEnumerer(ByRef urclRegCle As REGCLE)            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     urclRegCle        -                                                                             ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub mobjWinReg_CleEnumerer(ByRef urclRegCle As REGCLE)
   Dim ucfiClsFic As CLSFIC                                                                      ' Un enregistrement de la collection des extension de fichier
   Dim ucapClsApl As CLSAPL                                                                   ' Un enregistrement de la collection des reference application
   ' Je regarde si elle est valide
   Select Case Left$(urclRegCle.colCleSous(0), 1)
      Case vbNullString, "*"
         ' C'est la valeur default ou tous fichier
      Case "."
         ' C'est une extension valide
         ' Je recupere l'extension
         ucfiClsFic.strClsFicExtension = urclRegCle.colCleSous(0)
         ' Je regarde si l'extension a une reference
         Select Case LectureFic(ucfiClsFic.strClsFicExtension, ucfiClsFic.strClsFicProgID)
            Case True
               ' J'ai bien lu
               ' Je creer un evenement
               RaiseEvent FicEnumerer(ucfiClsFic)
            Case False
               ' Je n'ai pas lu
         End Select
      Case Else
         ' C'est une reference non valide
         ' Je recupere l'extension
         ucapClsApl.strClsAplProgID = urclRegCle.colCleSous(0)
         ' Je regarde si l'extension a une reference
         Select Case LectureApl(ucapClsApl.strClsAplProgID, ucapClsApl.strClsAplOuvre, ucapClsApl.strClsAplIcon)
            Case True
               ' J'ai bien lu
               ' Je creer un evenement
               RaiseEvent AplEnumerer(ucapClsApl)
            Case False
               ' Je n'ai pas lu
         End Select
   End Select
End Sub

