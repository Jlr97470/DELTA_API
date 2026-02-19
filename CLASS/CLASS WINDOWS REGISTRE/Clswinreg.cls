VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsWinReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2029
Attribute VB_Description = "PrjDllDeltaApi.clsWinReg"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'******************************************************************************
'***    Delta Copyright                                                             (31/05/2001)  ***
'******************************************************************************
'***    CLASS:                                                                                             ***
'***        clsRegWin                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      ----------    Pour La Gestion De La Base De Registre Windows  --------   ***
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
'***    Constante Qui Defini Le Nom De La Feuille                                        ***
'******************************************************************************
Private Const LOGFEUILLENOM = "clsRegWin"
Private Const LOGFEUILLETYPE = "CLASS"

'******************************************************************************
'***    Constante Qui Defini Le Controle Sur Les Cles Du Registre                ***
'******************************************************************************
Private Enum REGCONTROLENUM
   REGCONTROLREAD = &H20000
   REGCONTROLSYNCHRONIZE = &H100000
   REGCONTROLSTANDARDRIGHTSALL = &H1F0000
   REGCONTROLSTANDARDRIGHTSEXECUTE = (REGCONTROLREAD)
   REGCONTROLSTANDARDRIGHTSREAD = (REGCONTROLREAD)
   REGCONTROLSTANDARDRIGHTSREQUIRED = &HF0000
   REGCONTROLSTANDARDRIGHTSWRITE = (REGCONTROLREAD)
   REGCONTROLKEYCREATELINK = &H20
   REGCONTROLKEYCREATESUBKEY = &H4
   REGCONTROLKEYENUMERATESUBKEYS = &H8
   REGCONTROLKEYNOTIFY = &H10
   REGCONTROLKEYQUERYVALUE = &H1
   REGCONTROLKEYREAD = ((REGCONTROLSTANDARDRIGHTSREAD Or REGCONTROLKEYQUERYVALUE Or REGCONTROLKEYENUMERATESUBKEYS Or REGCONTROLKEYNOTIFY) And (Not REGCONTROLSYNCHRONIZE))
   REGCONTROLKEYEXECUTE = (REGCONTROLKEYREAD)
   REGCONTROLKEYSETVALUE = &H2
   REGCONTROLKEYWRITE = ((REGCONTROLSTANDARDRIGHTSWRITE Or REGCONTROLKEYSETVALUE Or REGCONTROLKEYCREATESUBKEY) And (Not REGCONTROLSYNCHRONIZE))
   REGCONTROLKEYALLACCESS = ((REGCONTROLSTANDARDRIGHTSALL Or REGCONTROLKEYQUERYVALUE Or REGCONTROLKEYSETVALUE Or REGCONTROLKEYCREATESUBKEY Or REGCONTROLKEYENUMERATESUBKEYS Or REGCONTROLKEYNOTIFY Or REGCONTROLKEYCREATELINK) And (Not REGCONTROLSYNCHRONIZE))
 End Enum

'******************************************************************************
'***    Constante Qui Defini Les Options Des Cles Du Registre                     ***
'******************************************************************************
Private Enum REGOPTIONENUM
   REGOPTIONNONVOLATILE
   REGOPTIONVOLATILE
   REGOPTIONBACKUPRESTORE
End Enum

'******************************************************************************
'***    Constante Qui Defini Le Resultat De La Fonction RegEcritureCle        ***
'******************************************************************************
Private Enum REGRESULTATENUM
   REGRESULTATREGCREATEDNEWKEY
   REGRESULTATREGOPENEDEXISTINGKEY
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Erreur D'Access A La Base De Registre      ***
'******************************************************************************
Private Enum REGERREURENUM
   REGERREURSUCCESS = 0&
   REGERREURINVALIDKEY = 2&
   REGERREURACCESSDENIED = 5&
   REGERREUROUTOFMEMORY = 14&
   REGERREURINVALIDPARAMETER = 87&
   REGERREURMOREDATA = 234&
   REGERREURNOMOREITEMS = 259&
   REGERREURBADDB = 1009&
   REGERREURBADKEY = 1010&
   REGERREURCANTOPEN = 1011&
   REGERREURCANTREAD = 1012&
   REGERREURCANTWRITE = 1013&
   REGERREURREGISTRYRECOVERED = 1014&
   REGERREURREGISTRYCORRUPT = 1015&
   REGERREURREGISTRYIOFAILED = 1016&
   REGERREURNOTREGISTRYFILE = 1017&
   REGERREURKEYDELETED = 1018&
   REGERREURNOLOGSPACE = 1019&
   REGERREURKEYHASCHILDREN = 1020&
   REGERREURCHILDMUSTBEVOLATILE = 1021&
   REGERREURRXACTINVALIDSTATE = 1369&
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Cles sectionss De La Base De Registre         ***
'******************************************************************************
Private Enum REGSECTIONENUM
   REGSECTIONCLASSESROOT = &H80000000
   REGSECTIONCURRENTCONFIG = &H80000005
   REGSECTIONCURRENTUSER = &H80000001
   REGSECTIONDYNDATA = &H80000006
   REGSECTIONLOCALMACHINE = &H80000002
   REGSECTIONPERFORMANCEDATA = &H80000004
   REGSECTIONUSERS = &H80000003
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Types De Valeur De La Base De Registre    ***
'******************************************************************************
Private Enum REGVALEURENUM
   REGVALEURNONE
   REGVALEURSZ
   REGVALEUREXPANDSZ
   REGVALEURBINARY
   REGVALEURDWORDLITTLEENDIAN
   REGVALEURDWORD
   REGVALEURDWORDBIGENDIAN
   REGVALEURLINK
   REGVALEURMULTISZ
 End Enum

'******************************************************************************
'***    Declaration De Type Priver                                                               ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGACCESSTYPE
    bytAccessRevision As Byte
    bytAccessBz1 As Byte
    intAccessTaille As Integer
    intAccessNombre As Integer
    intAccessBz2 As Integer
End Type

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGATTRIBUETYPE
    lngAttribueLongueur As Long
    lngAttribueDescription As Long
    lngAttribueEntete As Long
End Type

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGDESCRIPTIONTYPE
    bytDescriptionRevision As Byte
    bytDescriptionBz1 As Byte
    lngDescriptionControl As Long
    lngDescriptionOwner As Long
    lngDescriptionGroup As Long
    urcaDescriptionS As REGCONTROLENUM
    urcaDescriptionD As REGCONTROLENUM
End Type

'******************************************************************************
'***    Type Pour La Gestion  D'Une Erreur D'Access A La Base De Registre ***
'******************************************************************************
Private Type REGERREURTYPE
    lngErreurNumber As Long
    strErreurDescription As String
End Type

'******************************************************************************
'***    Type Pour La Gestion Du Dernier Access A La Base De Registre        ***
'******************************************************************************
Private Type REGDATETYPE
    lngDateBas As Long
    lngDateHaut As Long
End Type

'******************************************************************************
'***    Type Pour La Gestion Des Cles Primaires De La Base De Registre     ***
'******************************************************************************
Private Type REGSECTIONTYPE
   lngSectionCle As REGSECTIONENUM
   strSectionCle As String
End Type

'******************************************************************************
'***    Type Pour La Gestion Des Cles Secondaires De La Base De Registre  ***
'******************************************************************************
Private Type REGCLETYPE
   strCleIndex As String
   lngCleEntete As Long
   urdtCleDate As REGDATETYPE
   urcpCleSection As REGSECTIONTYPE
   blnCleSousEnumerer As Boolean
   colCleSous As Collection
   blnCleValeurEnumerer As Boolean
   colCleValeur As Collection
End Type

'******************************************************************************
'***    Type Pour La Gestion  D'Une Valeur De La Base De Registre              ***
'******************************************************************************
Private Type REGVALEURTYPE
    strValeurIndex As String
    strValeurNom As String
    lngValeurType As REGVALEURENUM
    vntValeurVba As Variant
End Type

'******************************************************************************
'***    Declaration De Fonction Priver Externe                                            ***
'******************************************************************************
    
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Ouvrire Une Cle Dans La Base De Registre                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                  - La Valeur De La Cle Primaire                ***
'***      strRegCle                        - La Valeur De La Cle Secondaire             ***
'***      lngRegOption                   - Les Options De La Cle                           ***
'***      lngRegControl                  - Le Control Sur La Cle                            ***
'***      lngRegCleEntete              - L'Entete De Cle Ouvert                          ***
'***    SORTIE:                                                                                           ***
'***      RegOuvreCle                   - Recupere Le Resultat De L'Operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function RegOuvreCle Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal lngRegSection As Long, ByVal strRegCle As String, ByVal lngRegOption As REGOPTIONENUM, ByVal lngRegControl As REGCONTROLENUM, ByRef lngRegCleEntete As Long) As REGERREURENUM

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Fermer Une Cle De La Base De Registre                                  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                - L'Entete De Cle Ouvert                        ***
'***    SORTIE:                                                                                           ***
'***      RegFermeCle                     - Recupere Le Resultat De L'Operation   ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function RegFermeCle Lib "advapi32.dll" Alias "RegCloseKey" (ByVal lngRegCleEntete As Long) As REGERREURENUM

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Ecrire Une Cle Dans La Base De Registre                                 ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                 - La Valeur De La Cle Primaire                 ***
'***      strRegCle                       - La Valeur De La Cle Secondaire              ***
'***      lngRegReserver                                                                              ***
'***      strRegClass                                                                                    ***
'***      lngRegOption                  - Les Options De La Cle                             ***
'***      lngRegControl                 - Le Control Sur La Cle                              ***
'***      lngRegSecuriter                                                                              ***
'***      lngRegCleEntete             - LEntete De Cle Ouvert                            ***
'***      lngRegDisposition                                                                           ***
'***    SORTIE:                                                                                           ***
'***      RegEcritureCle               - Recupere Le Resultat De L'Operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function RegEcritureCle Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal lngRegSection As REGSECTIONENUM, ByVal strRegCle As String, ByVal lngRegReserver As Long, ByVal strRegClass As String, ByVal lngRegOption As REGOPTIONENUM, ByVal lngRegControl As REGCONTROLENUM, ByVal lngRegSecuriter As Long, ByRef lngRegCleEntete As Long, ByRef lngRegDisposition As Long) As REGERREURENUM

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Supprimer Une Cle Dans La Base De Registre                          ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                 - La Valeur De La Cle Primaire                 ***
'***      strRegCle                        - La Valeur De La Cle Secondaire             ***
'***    SORTIE:                                                                                           ***
'***      RegSupprimeCle             - Recupere Le Resultat De L'Operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegSupprimeCle(lngRegSection, strRegCle)                                     ***
'******************************************************************************
Private Declare Function RegSupprimeCle Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngRegSection As Long, ByVal strRegCle As String) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lire Le Type Et La Longueur D'Une Valeur D'Une Cle                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete             - L'Entete De Cle Ouvert                            ***
'***      strRegValeurNom           - Le Nom De La Valeur                               ***
'***      lngRegReserver                                                                              ***
'***      lngRegValeurType          - Le Type De Valeur                                  ***
'***      anyNull                          - Valeur Null                                             ***
'***      lngRegValeurLongueur   - La Longueur De La Valeur                         ***
'***    SORTIE:                                                                                           ***
'***      RegLectureValeur           - Recupere Le Resultat De L'Operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Declare Function RegLectureValeur Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByRef lngRegValeurType As REGTYPECONSTANTES, ByRef anyNull As Any, ByRef lngRegValeurLongueur As Long) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lire La Valeur chaine de cle de la base de registre                  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                                - La l'entete de cle ouvert       ***
'***      strRegValeurNom                              - Le nom de la valeur              ***
'***      lngRegReserver                                                                              ***
'***      lngRegValeurType                              - Type de valeur                    ***
'***                                                                   CHAINE                            ***
'***      strRegValeur                                      - Valeur de la Cle                  ***
'***      lngRegValeurLongueur                        - Longueur de la valeur          ***
'***                                                                   de la Cle                           ***
'***    SORTIE:                                                                                           ***
'***      RegLectureValeurChaine                    - Recupere le status              ***
'***                                                                de  l'Operation                      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegLectureValeurChaine(lngRegCleEntete, strRegValeurNom, 0,      ***
'***     lngRegValeurType, strRegValeur, lngRegValeurLongueur)                 ***
'******************************************************************************
Private Declare Function RegLectureValeurChaine Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByRef lngRegValeurType As REGTYPECONSTANTES, ByVal strRegValeur As String, ByRef lngRegValeurLongueur As Long) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegEcritureValeur(ByVal lngRegCleEntete As Long,                        ***
'***      ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ***
'***      ByVal lngRegValeurType As REGTYPECONSTANTES                       ***
'***      ByVal lngRegValeur As Long,                                                          ***
'***      ByVal lngRegValeurLongueur As Long)                                            ***
'***      As REGERREURCONSTANTES                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire une valeur entier long de cle dans de la base de registre***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                                   - La l'entete de cle ouvert    ***
'***      strRegValeurNom                                 - Le nom de la valeur           ***
'***      lngRegReserver                                                                              ***
'***      lngRegValeurType                                - Type de valeur                  ***
'***      lngRegValeur                                        - Valeur de la de Cle           ***
'***      lngRegValeurLongueur                          - Longueur de la valeur        ***
'***                                                                    de la  Cle                          ***
'***    SORTIE:                                                                                            ***
'***      RegEcritureValeur                                 - Recupere le status           ***
'***                                                                    de  l'operation                   ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegEcritureValeur(lngRegCleEntete, strRegValeurNom, 0,               ***
'***     REG_BINARY, StrPtr(strValeur), 1)                                                   ***
'******************************************************************************
Private Declare Function RegEcritureValeur Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByVal lngRegValeurType As REGTYPECONSTANTES, ByVal lngRegValeur As Long, ByVal lngRegValeurLongueur As Long) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegEcritureValeurChaine                                                                ***
'***      (ByVal lngRegSection As Long,                                                       ***
'***      ByVal strRegValeurNom As String,                                                 ***
'***      ByVal lngRegReserver As Long,                                                      ***
'***      ByVal lngRegValeurType As REGTYPECONSTANTES,                     ***
'***      ByVal strRegValeur As String,                                                        ***
'***      ByVal lngRegValeurLongueur As Long)                                            ***
'***      As REGERREURCONSTANTES                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire une valeur chaine de cle dans de la base de registre       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                                - La l'entete de cle ouvert      ***
'***      strRegValeurNom                               - Le nom de la valeur             ***
'***      lngRegReserver                                                                              ***
'***      lngRegValeurType                               - Type de valeur                   ***
'***                                                                  CHAINE                             ***
'***      strRegValeur                                      - Valeur de la de Cle             ***
'***      lngRegValeurLongueur                        - Longueur de la valeur de la  ***
'***                                                                 Cle                                      ***
'***    SORTIE:                                                                                           ***
'***      RegEcritureValeurChaine                  - Recupere le status               ***
'***                                                                de  l'Operation                      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegEcritureValeurChaine(lngRegCleEntete, strRegValeurNom, 0,     ***
'***     REG_SZ, strRegValeur, Len(strRegValeur))                                      ***
'******************************************************************************
Private Declare Function RegEcritureValeurChaine Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByVal lngRegValeurType As REGTYPECONSTANTES, ByVal strRegValeur As String, ByVal lngRegValeurLongueur As Long) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegSupprimeValeur (ByVal lngRegCleEntete As Long,                     ***
'***      ByVal strRegValeurNom As String) As REGERREURCONSTANTES  ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour supprimer une cle dans la base de registre                               ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***      lngRegCleEntete                             - La l'entete de cle ouvert          ***
'***      strRegValeurNom                            - Le nom de la valeur                ***
'***    SORTIE:                                                                                           ***
'***      RegSupprimeValeur                         - Recupere le status  de           ***
'***                                                               l'Operation  externe               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegSupprimeValeur(lngRegCleEntete, strRegValeurNom)                 ***
'******************************************************************************
Private Declare Function RegSupprimeValeur Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     RegEnumereCle(ByVal lngRegCleEntete As Long,                             ***
'***     ByVal lngRegCleIndex As Long,ByVal strRegCle As String,               ***
'***     ByRef lngRegCleLongueur As Long,                                                  ***
'***     ByVal lngRegReserver As Long,                                                       ***
'***     ByVal strRegClass As String, ByRef lngRegClassLongueur As Long, ***
'***     ByRef urdhRegCleDateHeure As REGDATE)                                      ***
'***     As REGERREURCONSTANTES                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire une cle dans la base de registre                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                           - La l'entete de cle ouvert           ***
'***      lngRegCleIndex                             - L'index de sous cle                  ***
'***      strRegCle                                     - La sous cle                             ***
'***      lngRegCleLongueur                        - La longueur de la sous cle      ***
'***      lngRegReserver                                                                             ***
'***      lngRegClass                                                                                  ***
'***      lngRegClassLongueur                                                                     ***
'***      urdhRegCleDateHeure                    - La date  du dernier access       ***
'***    SORTIE:                                                                                           ***
'***      RegEnumereCle                             - Recupere le status  de             ***
'***                                                             l'Operation  externe                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     = RegEnumereCle(lngRegCleEntete, lngRegCleIndex,                       ***
'***     strRegCle,lngRegCleLongueur, 0, 0,                                                 ***
'***     0, ufilDernierAccess)                                                                       ***
'******************************************************************************
Private Declare Function RegEnumereCle Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal lngRegCleEntete As Long, ByVal lngRegCleIndex As Long, ByVal strRegCle As String, ByRef lngRegCleLongueur As Long, ByVal lngRegReserver As Long, ByVal lngRegClass As Long, ByRef lngRegClassLongueur As Long, ByRef urdhRegCleDateHeure As REGDATE) As Long

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     RegEnumereValeur Lib "advapi32.dll" Alias "RegEnumValueA"         ***
'***     (ByVal lngRegCleEntete As Long, ByVal lngRegValeurIndex As Long,***
'***     ByVal strRegValeurNom As String,                                                   ***
'***     ByRef lngRegValeurNomLongueur As Long,                                       ***
'***     ByVal lngRegReserver As Long,                                                       ***
'***     ByRef lngRegValeurType As REGTYPECONSTANTES,                       ***
'***     ByRef anyNull As Any, ByVal lngRegValeurLongueur As Long)          ***
'***     As Long                                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                           - La l'entete de cle ouvert           ***
'***      lngRegValeurIndex                        - L'index de la valeur                  ***
'***      strRegValeurNom                         - Le nom de la valeur de cle        ***
'***      lngRegValeurNomLongueur           - La longueur du nom de la valeur***
'***      lngRegReserver                                                                             ***
'***      lngRegValeurType                        - Type de valeur                         ***
'***      anyNull                                       - Valeur Null                               ***
'***      lngRegValeurLongueur                  - Longueur de la valeur                ***
'***                                                           de la  Cle                                  ***
'***    SORTIE:                                                                                           ***
'***      RegEnumereValeur                      - Recupere le status de                ***
'***                                                           l'Operation                                ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      = RegEnumereValeur(lngRegCleEntete, lngRegValeurIndex,             ***
'***      strRegValeurNom, lngRegValeurNomLongueur, 0,                            ***
'***      lngRegValeurType, 0, 0)                                                                 ***
'******************************************************************************
Private Declare Function RegEnumereValeur Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal lngRegCleEntete As Long, ByVal lngRegValeurIndex As Long, ByVal strRegValeurNom As String, ByRef lngRegValeurNomLongueur As Long, ByVal lngRegReserver As Long, ByRef lngRegValeurType As REGTYPECONSTANTES, ByVal anyNull As Any, ByRef lngRegValeurLongueur As Long) As Long
               
'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***      WinMemCopie(ByRef anyCopDestination As Any,                            ***
'***      ByRef anyCopSource As Any,                                                         ***
'***      ByVal lngCopLongueur As Long)                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour copier une partie de donner en memoire                                    ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     anyCopDestination                                 - La destination                  ***
'***     anyCopSource                                       - La source                         ***
'***     lngCopLongueur                                     - La longueur                      ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     WinMemCopie bytCopValeur, ByVal StrPtr(strCopValeur),1               ***
'******************************************************************************
Private Declare Sub WinMemCopie Lib "kernel32" Alias "RtlMoveMemory" (ByRef anyCopDestination As Any, ByRef anyCopSource As Any, ByVal lngCopLongueur As Long)
               
'******************************************************************************
'***    Declaration De Fonction Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CRegValeurVbaValeur(ByVal strRegValeur As String,                       ***
'***     ByVal lngRegValeurType As REGTYPECONSTANTES) As Variant      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertir une valeur de base de registre en valeur vb                ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     strRegValeur                     - Une valeur variant de base de registre  ***
'***     lngRegValeurType             - Le type de valeur  de base de registre   ***
'***     lngRegValeurLongueur       - La longueur de la valeur                        ***
'***    SORTIE:                                                                                           ***
'***     CRegValeurVbaValeur        - Retourne une valeur vb                         ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CRegValeurVbaValeur("TITI"+chr$(0),REG_SZ,5)                             ***
'******************************************************************************
Private Function CRegValeurVbaValeur(ByVal strRegValeur As String, ByVal lngRegValeurType As REGTYPECONSTANTES, ByVal lngRegValeurLongueur As Long) As Variant
Attribute CRegValeurVbaValeur.VB_Description = "******************************************************************************"
   Dim vntVbaValeur As Variant                                                                       ' La valeur variant
   Dim strVbaValeur  As String                                                                        ' La valeur chaine
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CRegValeurVbaValeur_Erreur
   ' Je regarde le type de valeur
   Select Case lngRegValeurType
      Case REG_SZ, REG_EXPAND_SZ
         ' C'est une valeur chaine
         ' Je remet en forme
         strVbaValeur = Mid(strRegValeur, 1, lngRegValeurLongueur - 1)
         ' Je regarde si c'est une date
         Select Case IsDate(strVbaValeur)
            Case True
               ' C'est une date je retourne une date
               ' Je reourne une date
               CRegValeurVbaValeur = CDate(strVbaValeur)
            Case False
               ' Je renvoie la valeur
               ' Je regarde si c'est un boolean
               Select Case strVbaValeur
                  Case "True"
                     ' C'est un boolean
                     ' Je le renvoie a l'user
                     CRegValeurVbaValeur = True
                  Case "False"
                     ' C'est un boolean
                     ' Je le renvoie a l'user
                     CRegValeurVbaValeur = False
                  Case Else
                     ' Ce n'est pas un boolean
                     ' Je renvoie une chaine
                     CRegValeurVbaValeur = strVbaValeur
               End Select
         End Select
      Case REG_MULTI_SZ
         ' C'est une valeur chaine multiple
         ' Je renvoie la valeur
         CRegValeurVbaValeur = CstrRegValeurstrVbaValeurTab(strRegValeur, lngRegValeurLongueur)
      Case REG_DWORD
         ' C'est une valeur entierlong
         ' Je renvoie la valeur
         CRegValeurVbaValeur = CstrRegValeurlngVbaValeur(strRegValeur)
      Case REG_BINARY
         ' C'est une valeur d'un tableau d'octet
         ' Je regarde la valeur de la longueur
         Select Case lngRegValeurLongueur
            Case 1
               ' C'est un octet
               ' Je renvoie un octet
               CRegValeurVbaValeur = CstrRegValeurbytVbaValeur(strRegValeur)
            Case 2
               ' c'est un entier
               ' Je renvoie un entier
               CRegValeurVbaValeur = CstrRegValeurintVbaValeur(strRegValeur)
            Case 4
               ' c'est un entier long
               ' Je renvoie le entier long
               CRegValeurVbaValeur = CstrRegValeurlngVbaValeur(strRegValeur)
            Case Else
               ' Je redimensionne le tableau d'octet
                ' Je renvoie le tableau d'octet
               CRegValeurVbaValeur = CstrRegValeurbytVbaValeurTab(strRegValeur, lngRegValeurLongueur)
         End Select
       Case Else
         ' c'est un type incconue
         ' Je renvoie une chaine vide
         CRegValeurVbaValeur = Empty
   End Select
   ' Fin
CRegValeurVbaValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CRegValeurVbaValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CRegValeurVbaValeur", _
      CStr(strRegValeur) & "," & CStr(lngRegValeurType) & "," & CStr(lngRegValeurLongueur), Err
   ' Je renvoie une chaine vide
   CRegValeurVbaValeur = Empty
   ' Je Continue
   Resume CRegValeurVbaValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CVbaValeurRegValeur(ByVal vntVbaValeur As Variant,                     ***
'***     ByVal lngRegValeurType As REGTYPECONSTANTES) As String        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertir une valeur  vb en valeur de base de registre               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     vntVbaValeur                     - Une valeur variant de visual basic         ***
'***     lngRegValeurType              - Le type de valeur  de base de registre   ***
'***     lngRegValeurLongueur        - La longueur de la valeur                       ***
'***    SORTIE:                                                                                           ***
'***     lngRegValeurType              - Le type de valeur  de base de registre   ***
'***     lngRegValeurLongueur        - La longueur de la valeur                       ***
'***     CVbaValeurRegValeur        - Retourne une valeur de base de registre***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CVbaValeurRegValeur("TITI",                                                         ***
'***     lngRegValeurType,lngRegValeurLongueur)                                        ***
'******************************************************************************
Private Function CVbaValeurRegValeur(ByVal vntVbaValeur As Variant, ByRef lngRegValeurType As REGTYPECONSTANTES, ByRef lngRegValeurLongueur As Long) As String
   Dim bytVbaValeurTableau() As Byte                                                            ' La valeur d'un tableau d'octet de visual basic
   Dim strVbaValeurTableau() As String                                                          ' La valeur tableau de chaine de visual basic
   Dim strRegValeur As String                                                                        ' La valeur chaine de la base de regsitre
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CVbaValeurRegValeur_Erreur
   ' Je regarde le type de variable
   Select Case VarType(vntVbaValeur)
      Case vbString
        ' C'est une chaine
        ' Je recupere la valeur registre chaine
        strRegValeur = vntVbaValeur & Chr(0)
        ' Je recupere le type de valeur de base de registre
        lngRegValeurType = REG_SZ
        ' Je recupere la longueur de la valuer de base de registre
        lngRegValeurLongueur = Len(strRegValeur)
        ' Je renvoie la valeur
        CVbaValeurRegValeur = strRegValeur
      Case vbDate, vbDecimal, vbDouble, vbSingle, vbCurrency, vbBoolean
         ' C'est un type date ...
         ' Je recupere la valeur registre chaine
         strRegValeur = Format(vntVbaValeur) & Chr(0)
         ' Je recupere le type de valeur de base de registre
         lngRegValeurType = REG_SZ
         ' Je recupere la longueur de la valuer de base de registre
         lngRegValeurLongueur = Len(strRegValeur)
         ' Je renvoie la valeur
         CVbaValeurRegValeur = strRegValeur
      Case vbByte
         ' C'est un octet
         ' Je recupere la valeur registre chaine
         strRegValeur = Chr(vntVbaValeur)
         ' Je recupere le type de valeur de base de registre
         lngRegValeurType = REG_BINARY
         ' Je recupere la longueur de la valuer de base de registre
         lngRegValeurLongueur = 1
         ' Je renvoie la valeur
         CVbaValeurRegValeur = strRegValeur
      Case vbLong, vbInteger
         ' c'est une valeur numerique
         ' Je defini l'espace chaine
         strRegValeur = ClngVbaValeurstrRegValeur(CLng(vntVbaValeur))
         ' Je recupere le type de valeur de base de registre
         lngRegValeurType = REG_DWORD
         ' Je recupere la longueur de la valuer de base de registre
         lngRegValeurLongueur = 4
         ' Je renvoie la valeur
         CVbaValeurRegValeur = strRegValeur
      Case (vbArray + vbString)
         ' C'est un tableau de chaine
         ' Je recupere la valeur variant
         strVbaValeurTableau = vntVbaValeur
         ' Je le convertie en valeur chaine
         strRegValeur = CstrVbaValeurTabstrRegValeur(strVbaValeurTableau)
         ' Je recupere le type de valeur de base de registre
         lngRegValeurType = REG_MULTI_SZ
         ' Je recupere la longueur de la valuer de base de registre
         lngRegValeurLongueur = Len(strRegValeur)
         ' Je renvoie la valeur
         CVbaValeurRegValeur = strRegValeur
      Case (vbArray + vbByte)
         ' C'est un tableau d'octet
         ' Je recupere la valeur variant
         bytVbaValeurTableau = vntVbaValeur
         ' Je le convertie en valeur chaine
         strRegValeur = CbytVbaValeurTabstrRegValeur(bytVbaValeurTableau)
         ' Je recupere le type de valeur de base de registre
         lngRegValeurType = REG_BINARY
         ' Je recupere la longueur de la valuer de base de registre
         lngRegValeurLongueur = Len(strRegValeur)
         ' Je renvoie la valeur
         CVbaValeurRegValeur = strRegValeur
      Case Else
         ' Je renvoie la valeur
         CVbaValeurRegValeur = Empty
   End Select
   ' Fin
CVbaValeurRegValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CVbaValeurRegValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, _
      "CVbaValeurRegValeur", CStr(vntVbaValeur) & "," & CStr(lngRegValeurType) & "," & CStr(lngRegValeurLongueur), Err
   ' Je renvoie une chaine vide
   CVbaValeurRegValeur = Empty
   ' Je Continue
   Resume CVbaValeurRegValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrRegValeurbytVbaValeur(ByVal strRegValeur As String) As Byte ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire une valeur chaine de la base en byte             ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        strRegValeur                            -  une valeur chaine de la base      ***
'***    SORTIE:                                                                                           ***
'***        CstrRegValeurbytVbaValeu      -  Retourne un byte                        ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrRegValeurbytVbaValeur(strRegValeur)                                     ***
'******************************************************************************
Private Function CstrRegValeurbytVbaValeur(ByVal strRegValeur As String) As Byte
   Dim bytVbaValeur As Byte                                                                           ' La valeur binaire
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrRegValeurbytVbaValeur_Erreur
   ' Je recopie les donnees
   bytVbaValeur = Asc(Mid(strRegValeur, 1, 1))
   ' Je renvoie le tableau d'octet
   CstrRegValeurbytVbaValeur = bytVbaValeur
   ' Fin
CstrRegValeurbytVbaValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrRegValeurbytVbaValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurbytVbaValeur", strRegValeur, Err
   ' Je renvoie une chaine vide
   CstrRegValeurbytVbaValeur = vbEmpty
   ' Je Continue
   Resume CstrRegValeurbytVbaValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrRegValeurintVbaValeur(ByVal strRegValeur As String)               ***
'***     As Integer                                                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire une valeur chaine de la base en entier           ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        strRegValeur                            - une valeur chaine de la base       ***
'***    SORTIE:                                                                                           ***
'***        CstrRegValeurintVbaValeur      - Retourne un entier                      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrRegValeurintVbaValeur(strRegValeur)                                     ***
'******************************************************************************
Private Function CstrRegValeurintVbaValeur(ByVal strRegValeur As String) As Integer
   Dim lngVbaValeurIndex As Long                                                                   ' La position dans la valeur
   Dim intVbaValeur As Long                                                                           ' La valeur entier long
   Dim bytVbaValeur As Byte                                                                           ' La valeur binaire
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrRegValeurintVbaValeur_Erreur
   ' Je fait defiler les lignes du tableau
   For lngVbaValeurIndex = 0 To 1
      ' Je recopie les donnees
      bytVbaValeur = Asc(Mid(strRegValeur, lngVbaValeurIndex + 1, 1))
      ' Je rajoute la ligne a la valeur chaine
      intVbaValeur = intVbaValeur + bytVbaValeur * (&H100 ^ lngVbaValeurIndex)
   ' Je ne sort que si j'ai fini de faire defiler les lignes
   Next lngVbaValeurIndex
   ' Je renvoie le tableau d'octet
   CstrRegValeurintVbaValeur = intVbaValeur
   ' Fin
CstrRegValeurintVbaValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrRegValeurintVbaValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurintVbaValeur", strRegValeur, Err
   ' Je renvoie une chaine vide
   CstrRegValeurintVbaValeur = vbEmpty
   ' Je Continue
   Resume CstrRegValeurintVbaValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrRegValeurlngVbaValeur(ByVal strRegValeur As String) As Long ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire une valeur chaine de la base en entier  long  ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        strRegValeur                         -  une valeur chaine de la base         ***
'***    SORTIE:                                                                                           ***
'***        CstrRegValeurlngVbaValeur   -  Retourne un entier long                ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrRegValeurlngVbaValeur(strRegValeur)                                     ***
'******************************************************************************
Private Function CstrRegValeurlngVbaValeur(ByVal strRegValeur As String) As Long
   Dim lngVbaValeurIndex As Long                                                                   ' La position dans la valeur
   Dim lngVbaValeur As Long                                                                           ' La valeur entier long
   Dim bytVbaValeur As Byte                                                                           ' La valeur binaire
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrRegValeurlngVbaValeur_Erreur
   ' Je fait defiler les lignes du tableau
   For lngVbaValeurIndex = 0 To 3
      ' Je recopie les donnees
      bytVbaValeur = Asc(Mid(strRegValeur, lngVbaValeurIndex + 1, 1))
      ' Je rajoute la ligne a la valeur chaine
      lngVbaValeur = lngVbaValeur + bytVbaValeur * (&H100 ^ lngVbaValeurIndex)
   ' Je ne sort que si j'ai fini de faire defiler les lignes
   Next lngVbaValeurIndex
   ' Je renvoie le tableau d'octet
   CstrRegValeurlngVbaValeur = lngVbaValeur
   ' Fin
CstrRegValeurlngVbaValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrRegValeurlngVbaValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurlngVbaValeur", strRegValeur, Err
   ' Je renvoie une chaine vide
   CstrRegValeurlngVbaValeur = vbEmpty
   ' Je Continue
   Resume CstrRegValeurlngVbaValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrRegValeurstrVbaValeurTab(ByVal strRegValeur As String          ***
'***     ByVal lngRegValeurLongueur As Long) As String()                             ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire une valeur chaine de la base en tableau        ***
'***     de chaine                                                                                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strRegValeur                               -  une valeur chaine de la base      ***
'***      lngRegValeurLongueur                 - La longueur de la valeur              ***
'***    SORTIE:                                                                                           ***
'***        CstrRegValeurstrVbaValeurTab - Retourne un tableau de chaine    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrRegValeurstrVbaValeurTab(strRegValeur)                                ***
'******************************************************************************
Private Function CstrRegValeurstrVbaValeurTab(ByVal strRegValeur As String, ByVal lngRegValeurLongueur As Long) As String()
    Dim lngRegValeurIndex As Long                                                                  ' La position dans la valeur chaine de la base
    Dim strVbaValeurTableau() As String                                                           ' La valeur tableau de chaine
    Dim intVbaValeurTableauMax As Integer                                                      ' Le nombre de ligne du tableau
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrRegValeurstrVbaValeurTab_Erreur
   ' Je met a zero le nombre de ligne
   intVbaValeurTableauMax = vbEmpty
   ' Je redimensionne le tableau
   ReDim strVbaValeurTableau(intVbaValeurTableauMax)
   ' Je fait defiler les caracteres de la chaine multiple
   For lngRegValeurIndex = 1 To lngRegValeurLongueur - 1
      ' Je regarde la valeur du caractere
      Select Case Mid(strRegValeur, lngRegValeurIndex, 1)
         Case Chr(0)
            ' C'est le caractere de fin de ligne
            ' Je passe a la ligne suivante
            intVbaValeurTableauMax = intVbaValeurTableauMax + 1
            ' Je redimensionne en gardent les donners
            ReDim Preserve strVbaValeurTableau(intVbaValeurTableauMax)
         Case Else
            ' C'est un autre caractere
            ' Je le met dans la ligne du tableau
            strVbaValeurTableau(intVbaValeurTableauMax) = strVbaValeurTableau(intVbaValeurTableauMax) & Mid(strRegValeur, lngRegValeurIndex, 1)
      End Select
   ' Je ne sort que si j'ai fini de faire defiler
   Next lngRegValeurIndex
   ' Je renvoie le tableau de chaine
   CstrRegValeurstrVbaValeurTab = strVbaValeurTableau
   ' Fin
CstrRegValeurstrVbaValeurTab_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrRegValeurstrVbaValeurTab_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurstrVbaValeurTab", strRegValeur, Err
   ' Je renvoie le tableau de chaine
   CstrRegValeurstrVbaValeurTab = strVbaValeurTableau
   ' Je Continue
   Resume CstrRegValeurstrVbaValeurTab_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrRegValeurbytVbaValeurTab(ByVal strRegValeur As String,        ***
'***     ByVal lngRegValeurLongueur As Long) As Byte()                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire une valeur chaine de la base                         ***
'***     en tableau de byte                                                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      strRegValeur                                 -  une valeur chaine de la base   ***
'***      lngRegValeurLongueur                   - La longueur de la valeur           ***
'***    SORTIE:                                                                                           ***
'***        CstrRegValeurbytVbaValeurTab   -  Retourne un tableau de byte    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrRegValeurbytVbaValeurTab(strRegValeur,                               ***
'***        lngRegValeurLongueur)                                                                 ***
'******************************************************************************
Private Function CstrRegValeurbytVbaValeurTab(ByVal strRegValeur As String, ByVal lngRegValeurLongueur As Long) As Byte()
   Dim lngRegValeurIndex As Long                                                                  ' La position dans la valeur chaine de la base
   Dim bytVbaValeurTableau() As Byte                                                            ' La valeur d'un tableau d'octet
   Dim intVbaValeurTableauMax As Integer                                                    ' Le nombre de ligne du tableau
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrRegValeurbytVbaValeurTab_Erreur
   ' Je redimensionne le tableau
   ReDim bytVbaValeurTableau(lngRegValeurLongueur - 1)
   ' Je fait defiler les caracteres de la chaine multiple
   For lngRegValeurIndex = 1 To lngRegValeurLongueur
      ' Je calcule l'index
      intVbaValeurTableauMax = lngRegValeurIndex - 1
      ' Je la valeur dans le tableau
      bytVbaValeurTableau(intVbaValeurTableauMax) = CByte(Asc(Mid(strRegValeur, lngRegValeurIndex, 1)))
   ' Je ne sort que si j'ai fini de faire defiler
   Next lngRegValeurIndex
   ' Je renvoie le tableau de chaine
   CstrRegValeurbytVbaValeurTab = bytVbaValeurTableau
   ' Fin
CstrRegValeurbytVbaValeurTab_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrRegValeurbytVbaValeurTab_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurbytVbaValeurTab", strRegValeur, Err
   ' Je renvoie le tableau d'octet
   CstrRegValeurbytVbaValeurTab = bytVbaValeurTableau
   ' Je Continue
   Resume CstrRegValeurbytVbaValeurTab_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     ClngVbaValeurstrRegValeur(ByVal lngVbaValeur As Long)                ***
'***     As String                                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion une valeur entier long en valeur chaine de base        ***
'***     de registre                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        lngVbaValeur                             -  La valeur entier long                 ***
'***    SORTIE:                                                                                           ***
'***        ClngVbaValeurstrRegValeur       -  Retourne une valeur chaine       ***
'***                                                            de la base                                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =ClngVbaValeurstrRegValeur(strValeurTableau)                               ***
'******************************************************************************
Private Function ClngVbaValeurstrRegValeur(ByVal lngVbaValeur As Long) As String
   Dim strRegValeur As String                                                                         ' La valeur chaine de la base de registre
   Dim lngVbaValeurIndex As Long                                                                   ' La position dans la valeur
   Dim bytVbaValeur As Byte                                                                           ' La valeur binaire
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo ClngVbaValeurstrRegValeur_Erreur
   ' Je fait defiler les lignes du tableau
   For lngVbaValeurIndex = 0 To 3
      ' Je recopie les donnees
      WinMemCopie bytVbaValeur, ByVal (VarPtr(lngVbaValeur) + lngVbaValeurIndex), 1
      ' Je rajoute la ligne a la valeur chaine
      strRegValeur = strRegValeur & Chr(bytVbaValeur)
   ' Je ne sort que si j'ai fini de faire defiler les lignes
   Next lngVbaValeurIndex
   ' Je renvoie la valeur chaine
   ClngVbaValeurstrRegValeur = strRegValeur
   ' Fin
ClngVbaValeurstrRegValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
ClngVbaValeurstrRegValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "ClngVbaValeurstrRegValeur", CStr(lngVbaValeur), Err
   ' Je renvoie une chaine vide
   ClngVbaValeurstrRegValeur = vbNullString
   ' Je Continue
   Resume ClngVbaValeurstrRegValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CstrVbaValeurTabstrRegValeur(strVbaValeurTableau() As String)     ***
'***     As String                                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire un tableau de chaine  en valeur chaine          ***
'***     de la base                                                                                        ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        strVbaValeurTableau                  -  Le tableau de chaine                ***
'***    SORTIE:                                                                                           ***
'***        CstrVbaValeurTabstrRegValeur  -  Retourne une valeur chaine       ***
'***                                                            de la base                                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CstrVbaValeurTabstrRegValeur(strValeurTableau)                          ***
'******************************************************************************
Private Function CstrVbaValeurTabstrRegValeur(strVbaValeurTableau() As String) As String
   Dim strRegValeur As String                                                                         ' La valeur chaine de la base de registre
   Dim vntVbaValeurLigne As Variant                                                               ' La valeur chaine d'une ligne du tableau
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CstrVbaValeurTabstrRegValeur_Erreur
   ' Je fait defiler les lignes du tableau
   For Each vntVbaValeurLigne In strVbaValeurTableau
      ' Je rajoute la ligne a la valeur chaine
       strRegValeur = strRegValeur & vntVbaValeurLigne & Chr(0)
   ' Je ne sort que si j'ai fini de faire defiler les lignes
   Next vntVbaValeurLigne
   ' Je renvoie la valeur chaine
   CstrVbaValeurTabstrRegValeur = strRegValeur
   ' Fin
CstrVbaValeurTabstrRegValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CstrVbaValeurTabstrRegValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CstrVbaValeurTabstrRegValeur", strVbaValeurTableau(0), Err
   ' Je renvoie une chaine vide
   CstrVbaValeurTabstrRegValeur = vbNullString
   ' Je Continue
   Resume CstrVbaValeurTabstrRegValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     CbytVbaValeurTabstrRegValeur(bytVbaValeurTableau() As Byte)      ***
'***     As String                                                                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour convertion binaire un tableau de octet en chaine                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        bytVbaValeurTableau                 - Le tableau de octet                    ***
'***    SORTIE:                                                                                           ***
'***        CbytVbaValeurTabstrRegValeur  -  Retourne une valeur chaine     ***
'***                                                              de la base                              ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =CbytVbaValeurTabstrRegValeur(bytVbaValeurTableau)                   ***
'******************************************************************************
Private Function CbytVbaValeurTabstrRegValeur(bytVbaValeurTableau() As Byte) As String
   Dim strRegValeur As String                                                                         ' La valeur chaine de la base de registre
   Dim vntVbaValeur As Variant                                                                       ' La valeur d'un octet
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo CbytVbaValeurTabstrRegValeur_Erreur
   ' Je fait defiler les lignes du tableau
   For Each vntVbaValeur In bytVbaValeurTableau
      ' Je rajoute la ligne a la valeur chaine
       strRegValeur = strRegValeur & Chr(vntVbaValeur)
   ' Je ne sort que si j'ai fini de faire defiler les lignes
   Next vntVbaValeur
   ' Je renvoie la valeur chaine
   CbytVbaValeurTabstrRegValeur = strRegValeur
   ' Fin
CbytVbaValeurTabstrRegValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
CbytVbaValeurTabstrRegValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "CbytVbaValeurTabstrRegValeur", CStr(bytVbaValeurTableau(0)), Err
   ' Je renvoie une chaine vide
   CbytVbaValeurTabstrRegValeur = vbNullString
   ' Je Continue
   Resume CbytVbaValeurTabstrRegValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     OuvreCle(ByVal lngRegSection As REGSECTIONCONSTANTES,        ***
'***        ByVal strRegCle As String) as Boolean                                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ouvrir une cle dans la base de registre                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        lngRegSection         -              La section                                      ***
'***        strRegCle                -              La cle                                            ***
'***    SORTIE:                                                                                           ***
'***        OuvreCle                 -              Retourne le resultat de l'operation  ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =OuvreCle()                                                                                     ***
'******************************************************************************
Public Function OuvreCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute OuvreCle.VB_HelpID = 2030
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim strVarType As String                                                                            ' Le type de variable
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je defini l'index de la cle
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' J'initialise le type de la variable
   strVarType = vbNullString
   ' Je regarde le type de la variable
   strVarType = TypeName(gcolRegCle(urclRegCle.strCleIndex))
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo OuvreCle_Erreur
   ' Je regarde le type de la  variable
   Select Case strVarType
      Case "REGCLE"
        ' C'est une valeur valide
         ' Je recupere l'enregistrement
         urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
      Case vbNullString
         ' Ce n'est pas une valeur valide
        ' Je defini la section prim
        urclRegCle.urcpClePrim.lngPrimSection = lngRegSection
        ' Je defini la cle prim
        urclRegCle.urcpClePrim.strPrimCle = strRegCle
        ' Je defini la collection de sous cle
        Set urclRegCle.colCleSous = New Collection
        ' Je defini la collection de valeur
        Set urclRegCle.colCleValeur = New Collection
        ' Je la rajoute dans la collection
        gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
   End Select
   ' Je regarde l'entete de cle
   Select Case urclRegCle.lngCleEntete
      Case vbEmpty
         ' La cle n'est pas ouvert
        ' Je l'enleve de la collection
        gcolRegCle.Remove urclRegCle.strCleIndex
         ' J'ouvre l'entete de cle dans la base de registre
         lngRegWinOperation = RegOuvreCle(urclRegCle.urcpClePrim.lngPrimSection, urclRegCle.urcpClePrim.strPrimCle, 0, REGCONTROLKEYALLACCESS, urclRegCle.lngCleEntete)
         ' Je regarde le statut de l'operation
         Select Case lngRegWinOperation
            Case REGERREURSUCCESS
                ' J'ai ouvert la cle
                ' Je la rajoute dans la collection
                gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
                ' Je renvoie la valeur
                OuvreCle = True
            Case Else
                ' Je n'ai pas ouvert la cle
                ' Je renvoie la valeur
                OuvreCle = False
                ' Je regarde le statut de l'operation
                Select Case lngRegWinOperation
                   Case REGERREURINVALIDKEY
                      ' La cle n'existe pas
                   Case Else
                      ' c'est une autre erreur
                      ' J'ajoute l'erreur dans le journal
                      AjouteRegErreur (lngRegWinOperation)
                End Select
         End Select
      Case Else
         ' La cle est ouverte
         ' Je renvoie la valeur
         OuvreCle = True
   End Select
    ' Fin
OuvreCle_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
OuvreCle_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "OuvreCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
   ' Je n'ai pas Supprimer la cle
   OuvreCle = False
   ' Je Continue
   Resume OuvreCle_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     FermeCle(ByVal lngRegSection As REGSECTIONCONSTANTES,       ***
'***     ByVal strRegCle As String) As Boolean                                            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour fermer une cle dans la base de registre                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        lngRegSection    -             La cle SECTION de la base de registre  ***
'***        strRegCle               -             La sous cle de la base de registre    ***
'***    SORTIE:                                                                                           ***
'***        FermeCle                 -              Retourne le resultat de l'operation  ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =FermeCle(HKEY_CLASSES_ROOT, ".XXX")                                     ***
'******************************************************************************
Public Function FermeCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute FermeCle.VB_HelpID = 2031
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim strVarType As String                                                                           ' Le type de variable
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je defini l'index de la cle
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' J'initialise le type de la variable
   strVarType = vbNullString
   ' Je regarde le type de la variable
   strVarType = TypeName(gcolRegCle(urclRegCle.strCleIndex))
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo FermeCle_Erreur
   ' Je regarde le type de la variable
   Select Case strVarType
      Case "REGCLE"
        ' C'est une valeur valide
        ' Je recupere l'enregistrement
        urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
        ' Je regarde l'entete de cle
        Select Case urclRegCle.lngCleEntete
          Case vbEmpty
             ' La cle n'est pas ouvert
             ' Je renvoie la valeur
             FermeCle = True
          Case Else
             ' La cle est ouverte
            ' Je l'enleve de la collection
            gcolRegCle.Remove urclRegCle.strCleIndex
             ' Je ferme l'entete de cle dans la base de registre
             lngRegWinOperation = RegFermeCle(urclRegCle.lngCleEntete)
             ' Je regarde le statut de l'operation
             Select Case lngRegWinOperation
                Case REGERREURSUCCESS
                    ' J'ai fermer la cle
                    ' Je me rappelle que la cle est fermer
                    urclRegCle.lngCleEntete = vbEmpty
                    ' Je la rajoute dans la collection
                    gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
                    ' Je renvoie la valeur
                    FermeCle = True
                Case Else
                    ' Je n'ai pas fermer la cle
                    ' Je renvoie la valeur
                    FermeCle = False
                    ' Je regarde le statut de l'operation
                    Select Case lngRegWinOperation
                       Case REGERREURINVALIDKEY
                          ' La cle n'existe pas
                       Case Else
                          ' c'est une autre erreur
                          ' J'ajoute l'erreur dans le journal
                          AjouteRegErreur (lngRegWinOperation)
                    End Select
             End Select
        End Select
      Case vbNullString
         ' Ce n'est pas une valeur valide
         ' Je renvoie la valeur
         FermeCle = False
   End Select
    ' Fin
FermeCle_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
FermeCle_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "FermeCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
   ' Je n'ai pas Supprimer la cle
   FermeCle = False
   ' Je Continue
   Resume FermeCle_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EnumereCle(ByVal lngRegSection As REGSECTIONCONSTANTES,    ***
'***     ByVal strRegCle As String) As Boolean                                            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour enumerer les sous cle de la base de registre                             ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                                 - La cle SECTION                      ***
'***      strRegCle                                          - La sous cle                          ***
'***    SORTIE:                                                                                           ***
'***      EnumereCle                                      - Le resultat de l'operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EnumereCle(HKEY_CLASSES_ROOT, ".XXX")                                 ***
'******************************************************************************
Public Function EnumereCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute EnumereCle.VB_HelpID = 2032
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegClePrim As REGCLE                                                                 ' Un enregistrement de la collection de cle
   Dim urclRegCleSous As REGCLE                                                                 ' Un enregistrement de la collection de cle
   Dim lngRegCleEnumLongueur As Long                                                         ' La longueur de la valeur de la sous cle
   Dim lngRegCleEnumIndex As Long                                                              ' L'index de la sous cle
   Dim strVarType As String                                                                            ' Le type de variable
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EnumereCle_Erreur
   ' J'ouvre la cle
   Select Case OuvreCle(lngRegSection, strRegCle)
    Case True
        ' J'ai bien ouvert la cle
        ' Je defini l'index de la cle
        urclRegClePrim.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
        ' Je recupere l'enregistrement
        urclRegClePrim = gcolRegCle(urclRegClePrim.strCleIndex)
        ' Je regarde si j'ai deja enumerer les sous cles
        Select Case urclRegClePrim.blnCleSousEnumerer
            Case True
                ' J'ai deja enumerer les sous cles
            Case False
                ' Je n'ai pas enumerer les sous cles
                ' Je l'enleve de la collection
                gcolRegCle.Remove urclRegClePrim.strCleIndex
                ' Je me met sur le premier index
                lngRegCleEnumIndex = vbEmpty
                ' Je repete
                Do
                    ' J'initialise la longuer de la valeur de class
                    lngRegCleEnumLongueur = 1024
                    ' Je defini la section prim
                    urclRegCleSous.urcpClePrim.lngPrimSection = urclRegClePrim.urcpClePrim.lngPrimSection
                    ' J'e defini la cle prim
                    urclRegCleSous.urcpClePrim.strPrimCle = Space$(lngRegCleEnumLongueur)
                    ' Je defini la collection de sous cle
                    Set urclRegCleSous.colCleSous = New Collection
                    ' Je defini la collection de valeur
                    Set urclRegCleSous.colCleValeur = New Collection
                    ' Je recupere la sous cle suivante
                    lngRegWinOperation = RegEnumereCle(urclRegClePrim.lngCleEntete, lngRegCleEnumIndex, urclRegCleSous.urcpClePrim.strPrimCle, lngRegCleEnumLongueur, vbEmpty, vbEmpty, vbEmpty, urclRegCleSous.urdtCleDate)
                    ' Je regarde la valeur de l'operation
                    Select Case lngRegWinOperation
                        Case REGERREURSUCCESS
                            ' J'ai enumerer une sous cle
                            ' J'ai  enumerer les cles
                            EnumereCle = True
                        Case REGERREURNOMOREITEMS
                            ' Je n'ai plus de sous cle a enumerer
                            ' J'ai  enumerer les cles
                            EnumereCle = True
                            ' Je sort
                            Exit Do
                        Case Else
                            ' J'ai une autre erreur
                            ' Je n'ai pas enumerer de cle
                            EnumereCle = False
                            ' J'ajoute l'erreur dans le journal
                            AjouteRegErreur (lngRegWinOperation)
                            ' Je sort
                            Exit Do
                    End Select
                    ' Je recupere la cle prim
                    urclRegCleSous.urcpClePrim.strPrimCle = urclRegClePrim.urcpClePrim.strPrimCle & "\" & Mid(urclRegCleSous.urcpClePrim.strPrimCle, 1, lngRegCleEnumLongueur)
                    ' Je defini l'index de la cle
                    urclRegCleSous.strCleIndex = Hex(urclRegCleSous.urcpClePrim.lngPrimSection) & "\" & urclRegCleSous.urcpClePrim.strPrimCle
                    ' Je desactive la gestion d'erreur
                    On Error Resume Next
                    ' J'initialise le type de la variable
                    strVarType = vbNullString
                    ' Je regarde le type de la variable
                    strVarType = TypeName(gcolRegCle(urclRegCleSous.strCleIndex))
                    ' En Cas D'Erreur Je Gere L'Erreur
                    On Error GoTo EnumereCle_Erreur
                    ' Je regarde le type de la variable
                    Select Case strVarType
                        Case "REGCLE"
                            ' C'est une valeur valide
                        Case vbNullString
                            ' Ce n'est pas une valeur valide
                            ' Je la rajoute la cle dans la collection
                            gcolRegCle.Add urclRegCleSous, urclRegCleSous.strCleIndex
                    End Select
                    ' Je la rajoute dans la collection de sous cle
                    urclRegClePrim.colCleSous.Add urclRegCleSous.urcpClePrim, urclRegCleSous.strCleIndex
                    ' Je passe a la suivante
                    lngRegCleEnumIndex = lngRegCleEnumIndex + 1
                Loop
                ' Je defini que les sous cle ont ete enumerer
                urclRegClePrim.blnCleSousEnumerer = True
                ' Je la rajoute dans la collection
                gcolRegCle.Add urclRegClePrim, urclRegClePrim.strCleIndex
        End Select
      Case False
         ' Je n'ai pas ouvert la cle
         ' Je n'ai pas enumerer de cle
         EnumereCle = False
   End Select
   ' Fin
EnumereCle_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EnumereCle_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "EnumereCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
   ' Je n'ai pas supprimer la valeur de la cle
   EnumereCle = False
   ' Je Continue
   Resume EnumereCle_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EcritureCle(ByVal lngRegSection As REGSECTIONCONSTANTES,     ***
'***     ByVal strRegCle As String) as boolean                                             ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ecrire une cle dans la base de registre                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        lngRegSection    -            La cle SECTION de la base de registre    ***
'***        strRegCle              -             La sous cle de la base de registre      ***
'***    SORTIE:                                                                                           ***
'***        EcritureCle            -              Retourne le resultat de l'operation    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EcritureCle(HKEY_CLASSES_ROOT, ".XXX")                                  ***
'******************************************************************************
Public Function EcritureCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute EcritureCle.VB_HelpID = 2033
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim strVarType As String                                                                           ' Le type de variable
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je defini l'index de la cle
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' J'initialise le type de la variable
   strVarType = vbNullString
   ' Je regarde le type de la variable
   strVarType = TypeName(gcolRegCle(urclRegCle.strCleIndex))
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureCle_Erreur
   ' Je regarde le type de la variable
   Select Case strVarType
      Case "REGCLE"
         ' C'est une valeur valide
         ' Je recupere l'enregistrement
         urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
      Case vbNullString
        ' Ce n'est pas une valeur valide
        ' Je defini la section prim
        urclRegCle.urcpClePrim.lngPrimSection = lngRegSection
        ' Je defini la cle prim
        urclRegCle.urcpClePrim.strPrimCle = strRegCle
        ' Je defini la collection de sous cle
        Set urclRegCle.colCleSous = New Collection
        ' Je defini la collection de valeur
        Set urclRegCle.colCleValeur = New Collection
        ' Je la rajoute dans la collection
        gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
   End Select
   ' Je regarde l'entete de cle
   Select Case urclRegCle.lngCleEntete
      Case vbEmpty
         ' La cle n'est pas ouvert
        ' Je l'enleve de la collection
        gcolRegCle.Remove urclRegCle.strCleIndex
         ' J'ecrit la cle dans la base de registre
         lngRegWinOperation = RegEcritureCle(urclRegCle.urcpClePrim.lngPrimSection, urclRegCle.urcpClePrim.strPrimCle, vbEmpty, REG_NONE, vbEmpty, REGCONTROLKEYALLACCESS, vbEmpty, urclRegCle.lngCleEntete, vbEmpty)
         ' Je regarde le statut de l'operation
         Select Case lngRegWinOperation
            Case REGERREURSUCCESS
                ' J'ai bien ecrit la cle
                ' Je la rajoute dans la collection
                gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
                ' Je renvoie la valeur
                EcritureCle = True
            Case Else
                ' Je n'ai pas ecrit la cle
                ' Je renvoie la valeur
                EcritureCle = False
                ' Je regarde le statut de l'operation
                Select Case lngRegWinOperation
                   Case REGERREURINVALIDKEY
                      ' La cle n'existe pas
                   Case Else
                      ' c'est une autre erreur
                      ' J'ajoute l'erreur dans le journal
                      AjouteRegErreur (lngRegWinOperation)
                End Select
         End Select
      Case Else
         ' La cle est ouverte
         ' Je renvoie la valeur
         EcritureCle = True
   End Select
    ' Fin
EcritureCle_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EcritureCle_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "EcritureCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
   ' Je n'ai pas Supprimer la cle
   EcritureCle = False
   ' Je Continue
   Resume EcritureCle_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     SupprimeCle(ByVal lngRegSection As REGSECTIONCONSTANTES,   ***
'***     ByVal strRegCle As String) As Boolean                                            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer une cle dans la base de registre                               ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        lngRegSection    -             La cle SECTION de la base de registre  ***
'***        strRegCle             -             La sous cle de la base de registre      ***
'***    SORTIE:                                                                                           ***
'***        SupprimeCle        -              Retourne le resultat de l'operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =SupprimeCle(HKEY_CLASSES_ROOT, ".XXX")                                ***
'******************************************************************************
Public Function SupprimeCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute SupprimeCle.VB_HelpID = 2034
   Dim blnRegWinOperation As Boolean                                                           ' Le statut de l'operation
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegClePrim As REGCLE                                                                 ' Un enregistrement de la collection de cle
   Dim vntRegCleSous As Variant                                                                    ' Un enregistrement de la collection de sous cle
   Dim strVarType As String                                                                            ' Le type de variable
   ' Je desactive la gestion d'erreur
   On Error Resume Next
   ' Je defini l'index de la cle
   urclRegClePrim.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' J'initialise le type de la variable
   strVarType = vbNullString
   ' Je regarde le type de la variable
   strVarType = TypeName(gcolRegCle(urclRegClePrim.strCleIndex))
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo SupprimeCle_Erreur
   ' Je regarde le type de la  variable
   Select Case strVarType
      Case "REGCLE"
         ' C'est une valeur valide
         ' Je recupere l'enregistrement
         urclRegClePrim = gcolRegCle(urclRegClePrim.strCleIndex)
      Case vbNullString
        ' Ce n'est pas une valeur valide
   End Select
   ' Je regarde si les sous cles sont enumerer
   Select Case urclRegClePrim.blnCleSousEnumerer
        Case True
            ' J'ai deja enumerer les cles
        Case False
            ' Je n'ai pas enumerer les cles
            ' Je regarde le status de l'enumeration des sous cles
            Select Case EnumereCle(lngRegSection, strRegCle)
                Case True
                    ' J'ai bien enumerer les sous cles
                    ' Je me rapelle que j'ai un enregistrement
                    strVarType = "REGCLE"
                    ' Je recupere l'enregistrement
                    urclRegClePrim = gcolRegCle(urclRegClePrim.strCleIndex)
                Case False
                    ' Je n'ai pas enumerer les sous cles
            End Select
   End Select
   ' Je regarde si j'ai un enregistrement
   Select Case strVarType
    Case "REGCLE"
        ' C'est une valeur valide
        ' Je regarde le statut de fermeture de la cle
        Select Case FermeCle(lngRegSection, strRegCle)
            Case True
                ' J'ai bien fermer la cle
                Select Case urclRegClePrim.colCleSous.Count
                    Case vbEmpty
                        ' Je n'ai pas des sous cle
                        ' Je m'en rapelle
                        blnRegWinOperation = True
                    Case Else
                        ' J'ai des sous cle
                        ' Je les fait defiler
                        For Each vntRegCleSous In urclRegClePrim.colCleSous
                            ' Je supprime la sous cle
                            blnRegWinOperation = SupprimeCle(vntRegCleSous.lngPrimSection, vntRegCleSous.strPrimCle)
                            ' Je regarde le statut de l'operation
                            Select Case blnRegWinOperation
                                Case True
                                    ' J'ai bien supprimer la sous cles
                                    ' Je l'enleve de la collection de sous cle
                                    urclRegClePrim.colCleSous.Remove Hex(vntRegCleSous.lngPrimSection) & "\" & vntRegCleSous.strPrimCle
                                Case False
                                    ' Je n'ai pas supprimer la sous cles
                                    ' Je sort
                                    Exit For
                            End Select
                        Next
                End Select
                ' Je regarde si j'ai bien supprimer les sous cles
                Select Case blnRegWinOperation
                    Case True
                        ' J'ai bien supprimer les sous cles
                        ' Je supprime la cle
                        lngRegWinOperation = RegSupprimeCle(urclRegClePrim.urcpClePrim.lngPrimSection, urclRegClePrim.urcpClePrim.strPrimCle)
                        ' Je regarde le statut de l'operation
                        Select Case lngRegWinOperation
                           Case REGERREURSUCCESS
                               ' J'ai bien supprimer la cle
                               gcolRegCle.Remove urclRegClePrim.strCleIndex
                               ' Je renvoie la valeur
                               SupprimeCle = True
                           Case Else
                               ' Je n'ai pas supprimer la cle
                               ' Je renvoie la valeur
                               SupprimeCle = False
                                ' J'ajoute l'erreur dans le journal
                                AjouteRegErreur (lngRegWinOperation)
                        End Select
                    Case False
                        ' Je n'ai pas supprimer les sous cles
                        ' Je renvoie la valeur
                        SupprimeCle = False
                End Select
            Case False
                ' J'ai bien fermer la cle
                ' Je renvoie la valeur
                SupprimeCle = False
        End Select
      Case vbNullString
         ' Ce n'est pas une valeur valide
         ' Je renvoie la valeur
         SupprimeCle = False
   End Select
   ' Fin
SupprimeCle_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
SupprimeCle_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "SupprimeCle", CStr(lngRegSection) & "," & strRegCle, Err
   ' Je n'ai pas Supprimer la cle
   SupprimeCle = False
   ' Je Continue
   Resume SupprimeCle_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     LectureValeur(ByVal lngRegSection As REGSECTIONCONSTANTES, ***
'***     ByVal strRegCle As String, ByVal strRegValeurNom As String)        ***
'***      As Variant                                                                                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour lire une valeur de la base de registre                                        ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                         - La section prim                             ***
'***      strRegCle                                - La cle prim                                   ***
'***      strRegValeurNom                    - Le nom de la valeur                      ***
'***      vntVbaValeur                          - La valeur de la cle                         ***
'***    SORTIE:                                                                                           ***
'***      vntVbaValeur                          - La valeur de la cle                          ***
'***      LectureValeur                          - Renvoie l'etat de l'operation           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =LectureValeur(HKEY_CLASSES_ROOT, ".XXX", "STR",vntValeur)    ***
'******************************************************************************
Public Function LectureValeur(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String, ByVal strRegValeurNom As String, ByRef vntVbaValeur As Variant) As Boolean
Attribute LectureValeur.VB_HelpID = 2035
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim urvaRegValeur As REGVALEUR                                                            ' Un enregistrement de la collection de valeur
   Dim strRegValeur As String                                                                         ' La valeur de la base de registre
   Dim lngRegValeurLongueur As Long                                                              ' La longueur de la valeur de la base de registre
   Dim strVarType As String                                                                            ' Le type de variable
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo LectureValeur_Erreur
   ' Je defini la cle de l'enregistrement courant de la collection
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' Je regarde si j'ai bien ouvert l'entete de cle
   Select Case OuvreCle(lngRegSection, strRegCle)
      Case True
        ' J'ai bien ouvert la cle
        ' Je recupere l'enregistrement
        urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
        ' Je l'enleve de la la collection
        gcolRegCle.Remove urclRegCle.strCleIndex
        ' Je defini l'index de la valeur
        urvaRegValeur.strValeurIndex = strRegValeurNom
        ' Je desactive la gestion d'erreur
        On Error Resume Next
        ' J'initialise le type de la variable
        strVarType = vbNullString
        ' Je regarde le type de la variable
        strVarType = TypeName(urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex))
        ' En Cas D'Erreur Je Gere L'Erreur
        On Error GoTo LectureValeur_Erreur
        ' Je regarde le type de variable
        Select Case strVarType
           Case "REGVALEUR"
                ' C'est une valeur valide
                ' Je recupere l'enregistrement si elle  existe
                urvaRegValeur = urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex)
                ' Je defini la valeur
                vntVbaValeur = urvaRegValeur.vntValeurVba
                ' Je renvoie j'ai lu la valeur
                LectureValeur = True
           Case vbNullString
              ' Ce n'est pas une valeur valide
              ' Je defini le nom de la valeur
              urvaRegValeur.strValeurNom = strRegValeurNom
              ' Je recupere la longueur de la valeur et le type de valeur
              lngRegWinOperation = RegLectureValeur(urclRegCle.lngCleEntete, urvaRegValeur.strValeurNom, vbEmpty, urvaRegValeur.lngValeurType, Null, lngRegValeurLongueur)
              ' Je regarde si j'ai recuperer la longueur de la valeur
              Select Case lngRegValeurLongueur
                 Case vbEmpty
                    ' Je n'ai pas recuperer la longueur de la valeur
                    ' Je renvoie que je n'ai pas lu la valeur
                    LectureValeur = False
                 Case Else
                    ' J'ai bien recuperer la longueur de la valeur
                    ' Je rajoute 2 a la longueur pour les caracteres de fin chaine
                    lngRegValeurLongueur = lngRegValeurLongueur + 2
                    ' J'initialise la valeur chaine tampon
                    strRegValeur = Space$(lngRegValeurLongueur)
                    ' Je recupere la valeur chaine
                    lngRegWinOperation = RegLectureValeurChaine(urclRegCle.lngCleEntete, urvaRegValeur.strValeurNom, vbEmpty, urvaRegValeur.lngValeurType, strRegValeur, lngRegValeurLongueur)
                    ' Je regarde le statut de l'operation
                    Select Case lngRegWinOperation
                       Case REGERREURSUCCESS
                            ' J'ai bien supprimer la cle
                            ' Je renvoie que j'ai lu la valeur
                            LectureValeur = True
                       Case Else
                            ' Je n'ai pas supprimer la cle
                            ' Je renvoie que je n'ai pas lu la valeur
                            LectureValeur = False
                            ' J'ajoute l'erreur dans le journal
                            AjouteRegErreur (lngRegWinOperation)
                    End Select
                    ' Je recupere la valeur vb
                    urvaRegValeur.vntValeurVba = CRegValeurVbaValeur(strRegValeur, urvaRegValeur.lngValeurType, lngRegValeurLongueur)
                    ' Je la rajoute dans la collection
                    urclRegCle.colCleValeur.Add urvaRegValeur, urvaRegValeur.strValeurIndex
                    ' Je renvoie la valeur
                    vntVbaValeur = urvaRegValeur.vntValeurVba
              End Select
        End Select
        ' Je l'ajoute a la collection
        gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
      Case False
         ' Je n'ai pas ouvert l'entete de cle
         ' Je renvoie que je n'ai pas lu la valeur
         LectureValeur = False
   End Select
    ' Fin
LectureValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
LectureValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "LectureValeur", CStr(lngRegSection) & "," & CStr(strRegCle) & "," & CStr(strRegValeurNom), Err
   ' Je renvoie la valeur
   LectureValeur = False
   ' Je Continue
   Resume LectureValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EcritureValeur(ByVal lngRegSection As REGSECTIONCONSTANTES,***
'***     ByVal strRegCle As String, ByVal strRegValeurNom As String,         ***
'***     ByVal vntVbaValeur As Variant) As Boolean                                     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ecrire une valeur de la base de registre                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                                   - La section prim                    ***
'***      strRegCle                                          - La cle prim                         ***
'***      strRegValeurNom                              - Le nom de la valeur              ***
'***      vntVbaValeur                                    - La valeur                              ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EcritureValeur(HKEY_CLASSES_ROOT, ".XXX", "STR","COUCOU") ***
'******************************************************************************
Public Function EcritureValeur(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String, ByVal strRegValeurNom As String, ByVal vntVbaValeur As Variant) As Boolean
Attribute EcritureValeur.VB_HelpID = 2036
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim urvaRegValeur As REGVALEUR                                                            ' Un enregistrement de la collection de valeur
   Dim strRegValeur As String                                                                         ' La valeur de la base de registre
   Dim lngRegValeurLongueur As Long                                                              ' La longueur de la valeur de la base de registre
   Dim strVarType As String                                                                            ' Le type de variable
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EcritureValeur_Erreur
   ' Je defini l'index de la cle
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' Je regarde si j'ai bien ouvert la cle
   Select Case OuvreCle(lngRegSection, strRegCle)
      Case True
         ' J'ai bien ouvert la cle
        ' Je recupere l'enregistrement
        urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
        ' Je l'enleve de la la collection
        gcolRegCle.Remove urclRegCle.strCleIndex
        ' Je defini l'index de la valeur
        urvaRegValeur.strValeurIndex = strRegValeurNom
        ' Je desactive la gestion d'erreur
        On Error Resume Next
        ' J'initialise le type de la variable
        strVarType = vbNullString
        ' Je regarde le type de la variable
        strVarType = TypeName(urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex))
        ' En Cas D'Erreur Je Gere L'Erreur
        On Error GoTo EcritureValeur_Erreur
        ' Je regarde le type de variable
        Select Case strVarType
           Case "REGVALEUR"
                ' C'est une valeur valide
                ' Je recupere l'enregistrement si elle  existe
                urvaRegValeur = urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex)
                ' Je l'enleve si elle existe dans la collection
                urclRegCle.colCleValeur.Remove urvaRegValeur.strValeurIndex
            Case vbNullString
                ' Ce n'est pas une valeur valide
                ' Je defini le nom de la valeur
                urvaRegValeur.strValeurNom = strRegValeurNom
         End Select
         ' Je defini la valeur
         urvaRegValeur.vntValeurVba = vntVbaValeur
         ' Je converti la valeur
         strRegValeur = CVbaValeurRegValeur(urvaRegValeur.vntValeurVba, urvaRegValeur.lngValeurType, lngRegValeurLongueur)
         ' J'ecrit la valeur
         lngRegWinOperation = RegEcritureValeurChaine(urclRegCle.lngCleEntete, urvaRegValeur.strValeurNom, vbEmpty, urvaRegValeur.lngValeurType, strRegValeur, lngRegValeurLongueur)
         ' Je regarde le statut de l'operation
         Select Case lngRegWinOperation
            Case REGERREURSUCCESS
                ' J'ai bien ecrit la valeur de la cle
                ' Je l'enleve si elle existe dans la collection
                urclRegCle.colCleValeur.Add urvaRegValeur, urvaRegValeur.strValeurNom
                'Je renvoie que la valeur n'a pas ete ecrit
                EcritureValeur = True
            Case Else
               ' Je n'ai pas ecrit la valeur de la cle
               'Je renvoie que la valeur n'a pas ete ecrit
               EcritureValeur = False
               ' J'ajoute l'erreur dans le journal
               AjouteRegErreur (lngRegWinOperation)
         End Select
        ' Je l'enleve de la la collection
        gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
      Case False
         ' Je n'ai pas ouvert l'entete de cle
         ' Je n'ai pas ecrit la valeur de la cle
         EcritureValeur = False
   End Select
    ' Fin
EcritureValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EcritureValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "EcritureValeur", CStr(lngRegSection) & "," & CStr(strRegCle) & "," & CStr(strRegValeurNom) & "," & CStr(vntVbaValeur), Err
   'Je renvoie que la valeur n'a pas ete ecrit
  EcritureValeur = False
   ' Je Continue
   Resume EcritureValeur_Exit
   ' Fin
End Function
  
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     SupprimeValeur                                                                               ***
'***     (ByVal lngRegSection As REGSECTIONCONSTANTES,                      ***
'***     ByVal strRegCle As String, ByVal strValeurNom As String,              ***
'***     ) As Boolean                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour supprimer une valeur de la base de registre                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                              - La section prim                         ***
'***      strRegCle                                    - La cle prim                                ***
'***      strRegValeurNom                        - Le nom de la valeur                    ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =SupprimeValeur(HKEY_CLASSES_ROOT, ".XXX", "STR")                ***
'******************************************************************************
Public Function SupprimeValeur(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String, ByVal strRegValeurNom As String) As Boolean
Attribute SupprimeValeur.VB_HelpID = 2037
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim urvaRegValeur As REGVALEUR                                                            ' Un enregistrement de la collection de valeur
   Dim strVarType As String                                                                            ' Le type de variable
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo SupprimeValeur_Erreur
   ' Je defini la cle de l'enregistrement courant de la collection
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' Je regarde si j'ai bien ouvert l'entete de cle
   Select Case OuvreCle(lngRegSection, strRegCle)
      Case True
            ' J'ai bien ouvert l'entete de cle
            ' Je recupere l'enregistrement
            urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
            ' Je l'enleve de la la collection
            gcolRegCle.Remove urclRegCle.strCleIndex
            ' Je defini l'index de la valeur
            urvaRegValeur.strValeurIndex = strRegValeurNom
            ' Je desactive la gestion d'erreur
            On Error Resume Next
            ' J'initialise le type de la variable
            strVarType = vbNullString
            ' Je regarde le type de la variable
            strVarType = TypeName(urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex))
            ' En Cas D'Erreur Je Gere L'Erreur
            On Error GoTo SupprimeValeur_Erreur
            ' Je regarde le type de variable
            Select Case strVarType
                Case "REGVALEUR"
                     ' C'est une valeur valide
                     ' Je recupere l'enregistrement si elle  existe
                     urvaRegValeur = urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex)
                     ' Je supprime la valeur de cle
                     lngRegWinOperation = RegSupprimeValeur(urclRegCle.lngCleEntete, urvaRegValeur.strValeurNom)
                     ' Je regarde le statut de l'operation
                     Select Case lngRegWinOperation
                        Case REGERREURSUCCESS
                           ' J'ai bien supprimer la valeur de la cle
                           ' Je l'enleve si elle existe dans la collection
                           urclRegCle.colCleValeur.Remove urvaRegValeur.strValeurIndex
                           ' J'ai bien supprimer la valeur de la cle
                           SupprimeValeur = True
                        Case Else
                           ' J'ai bien supprimer la valeur de la cle
                           ' Je n'ai pas supprimer la valeur de la cle
                           SupprimeValeur = False
                           ' J'ajoute l'erreur dans le journal
                           AjouteRegErreur (lngRegWinOperation)
                     End Select
                 Case False
                     ' Ce n'est pas une valeur valide
                     ' Je n'ai pas supprimer la valeur de la cle
                     SupprimeValeur = False
            End Select
            ' Je l'ajoute a la collection
            gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
      Case False
        ' Je n'ai pas ouvert l'entete de cle
        ' Je n'ai pas supprimer la valeur de la cle
        SupprimeValeur = False
   End Select
   ' Fin
SupprimeValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
SupprimeValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "SupprimeValeur", CStr(lngRegSection) & "," & strRegCle & "," & strRegValeurNom, Err
   ' Je n'ai pas supprimer la valeur de la cle
   SupprimeValeur = False
   ' Je Continue
   Resume SupprimeValeur_Exit
   ' Fin
End Function
  
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     EnumereValeur                                                                                ***
'***     (ByVal lngRegSection As REGSECTIONCONSTANTES,                      ***
'***     ByVal strRegCle As String) As Boolean                                            ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour enumerer les valeurs de la base de registre                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                                  - La section prim                     ***
'***      strRegCle                                         - La cle prim                           ***
'***    SORTIE:                                                                                           ***
'***      EnumereValeur                                 - Le resultat de l'operation      ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     =EnumereValeur(HKEY_CLASSES_ROOT, ".XXX")                             ***
'******************************************************************************
Public Function EnumereValeur(ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As Boolean
Attribute EnumereValeur.VB_HelpID = 2038
   Dim lngRegWinOperation As Long                                                                ' Le statut de l'operation
   Dim urclRegCle As REGCLE                                                                        ' Un enregistrement de la collection de cle
   Dim urvaRegValeur As REGVALEUR                                                            ' Un enregistrement de la collection de valeur
   Dim strRegValeur As String                                                                         ' La valeur de la base de registre
   Dim lngRegValeurIndex As Long                                                                   ' L'index de la valeur
   Dim lngRegValeurNomLongueur As Long                                                      ' La longueur du nom de la valeur
   Dim lngRegValeurLongueur As Long                                                              ' La longueur de la valeur de la base de registre
   Dim strVarType As String                                                                            ' Le type de variable
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo EnumereValeur_Erreur
   ' Je defini la cle de l'enregistrement courant de la collection
   urclRegCle.strCleIndex = Hex(lngRegSection) & "\" & strRegCle
   ' Je regarde le statut de l'operation
   Select Case OuvreCle(lngRegSection, strRegCle)
      Case True
            ' J'ai bien ouvert la cle
            ' Je recupere la cle courante
            urclRegCle = gcolRegCle(urclRegCle.strCleIndex)
            ' Je regarde si j'ai deja enumerer les valeur
            Select Case urclRegCle.blnCleValeurEnumerer
                Case True
                    ' J'ai deja enumerer les valeurs
                Case False
                    ' Je n'ai pas enumerer les valeurs
                    ' Je l'enleve de la la collection
                    gcolRegCle.Remove urclRegCle.strCleIndex
                    ' Je repete
                    Do
                        ' J'initialise la longuer de la valeur du nom
                        lngRegValeurNomLongueur = 1024
                        ' J'initialise la valeur du nom
                        urvaRegValeur.strValeurNom = Space$(lngRegValeurNomLongueur)
                        ' Je recupere la valeur suivante
                        lngRegWinOperation = RegEnumereValeur(urclRegCle.lngCleEntete, lngRegValeurIndex, urvaRegValeur.strValeurNom, lngRegValeurNomLongueur, vbEmpty, urvaRegValeur.lngValeurType, vbEmpty, lngRegValeurLongueur)
                        ' Je regarde si j'ai recuperer la longueur de la valeur
                        Select Case lngRegWinOperation
                            Case REGERREURSUCCESS
                                ' J'ai enumerer une sous cle
                                ' J'ai  enumerer les cles
                                EnumereValeur = True
                            Case REGERREURNOMOREITEMS
                                ' Je n'ai plus de sous cle a enumerer
                                ' J'ai  enumerer les cles
                                EnumereValeur = True
                                ' Je sort
                                Exit Do
                            Case Else
                                ' J'ai une autre erreur
                                ' Je n'ai pas enumerer de cle
                                EnumereValeur = False
                                ' J'ajoute l'erreur dans le journal
                                AjouteRegErreur (lngRegWinOperation)
                                ' Je sort
                                Exit Do
                        End Select
                        ' Je recupere la cle prim
                        urvaRegValeur.strValeurNom = Mid(urvaRegValeur.strValeurNom, 1, lngRegValeurNomLongueur)
                        ' Je defini l'index de la valeur
                        urvaRegValeur.strValeurIndex = urvaRegValeur.strValeurNom
                        ' Je desactive la gestion d'erreur
                        On Error Resume Next
                        ' J'initialise le type de la variable
                        strVarType = vbNullString
                        ' Je regarde le type de la variable
                        strVarType = TypeName(urclRegCle.colCleValeur(urvaRegValeur.strValeurIndex))
                        ' En Cas D'Erreur Je Gere L'Erreur
                        On Error GoTo EnumereValeur_Erreur
                        ' Je regarde le type de variable
                        Select Case strVarType
                           Case "REGVALEUR"
                                ' C'est une valeur valide
                            Case vbNullString
                                ' Ce n'est pas une valeur valide
                                ' Je rajoute 2 a la longueur pour les caracteres de fin chaine
                                lngRegValeurLongueur = lngRegValeurLongueur + 2
                                ' J'initialise la valeur chaine tampon
                                strRegValeur = Space$(lngRegValeurLongueur)
                                ' Je recupere la valeur chaine
                                lngRegWinOperation = RegLectureValeurChaine(urclRegCle.lngCleEntete, urvaRegValeur.strValeurNom, vbEmpty, urvaRegValeur.lngValeurType, strRegValeur, lngRegValeurLongueur)
                                ' Je regarde le statut de l'operation
                                Select Case lngRegWinOperation
                                   Case REGERREURSUCCESS
                                        ' J'ai bien lu la valeur
                                        ' J'ai  enumerer les cles
                                        EnumereValeur = True
                                   Case Else
                                        ' Je n'ai pas lu la valeur
                                        ' Je n'ai pas enumerer de cle
                                        EnumereValeur = False
                                        ' J'ajoute l'erreur dans le journal
                                        AjouteRegErreur (lngRegWinOperation)
                                        ' Je sort
                                        Exit Do
                                End Select
                                ' Je recupere la valeur vb
                                urvaRegValeur.vntValeurVba = CRegValeurVbaValeur(strRegValeur, urvaRegValeur.lngValeurType, lngRegValeurLongueur)
                                ' Je la rajoute dans la collection
                                urclRegCle.colCleValeur.Add urvaRegValeur, urvaRegValeur.strValeurIndex
                        End Select
                        ' Je passe a la suivante
                        lngRegValeurIndex = lngRegValeurIndex + 1
                    Loop
                    ' Je defini que les sous cle ont ete enumerer
                    urclRegCle.blnCleValeurEnumerer = True
                    ' Je la rajoute a la collection
                    gcolRegCle.Add urclRegCle, urclRegCle.strCleIndex
            End Select
      Case False
            ' Je n'ai pas ouvert la cle
            ' Je retourne que je n'ai pas enumere de cle
            EnumereValeur = False
   End Select
   ' Fin
EnumereValeur_Exit:
   ' Je Sort De La Procedure
   Exit Function
   ' Fin
EnumereValeur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEFONCTION, "EnumereValeur", CStr(lngRegSection) & "," & strRegCle, Err
   ' Je n'ai pas supprimer la valeur de la cle
   EnumereValeur = False
   ' Je Continue
   Resume EnumereValeur_Exit
   ' Fin
End Function

'******************************************************************************
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***     AjouteRegErreur(ByVal lngRegWinOperation As Long)                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour ecrire une erreur d'access a la base de registre                         ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***     lngRegWinOperation       - La valeur de l'operation                            ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     AjouteRegErreur(ERROR_INVALID_PARAMETER)                             ***
'******************************************************************************
Private Sub AjouteRegErreur(ByVal lngRegWinOperation As Long)
   Dim urerRegErreur As REGERREUR                                                           ' Un enregistrement d'erreur
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo AjouteRegErreur_Erreur
   ' Je defini le  numero d'erreur
   urerRegErreur.lngErreurNumber = lngRegWinOperation
   ' Je regarde le numero d'erreur
   Select Case lngRegWinOperation
      Case REGERREURACCESSDENIED
         urerRegErreur.strErreurDescription = "Access Denier"
      Case REGERREUROUTOFMEMORY
         urerRegErreur.strErreurDescription = "Pas assez de memoire"
      Case REGERREURINVALIDPARAMETER
         urerRegErreur.strErreurDescription = "Parametre invalide"
      Case REGERREURMOREDATA
         urerRegErreur.strErreurDescription = "Plus de donner"
      Case REGERREURNOMOREITEMS
         urerRegErreur.strErreurDescription = "Plus d'elements"
      Case REGERREURBADDB
         urerRegErreur.strErreurDescription = "The configuration registry database is corrupt."
      Case REGERREURBADKEY
         urerRegErreur.strErreurDescription = "The configuration registry key is invalid."
      Case REGERREURCANTOPEN
         urerRegErreur.strErreurDescription = "The configuration registry key could not be opened."
      Case REGERREURCANTREAD
         urerRegErreur.strErreurDescription = "The configuration registry key could not be read."
      Case REGERREURCANTWRITE
         urerRegErreur.strErreurDescription = "The configuration registry key could not be written."
      Case REGERREURREGISTRYRECOVERED
         urerRegErreur.strErreurDescription = "One of the files in the Registry database had to be recovered " & _
                   "by use of a log or alternate copy. The recovery was successful."
      Case REGERREURREGISTRYCORRUPT
         urerRegErreur.strErreurDescription = "The Registry is corrupt. The structure of one of the files that contains " & _
                   "Registry data is corrupt, or the system's image of the file in memory " & _
                   "is corrupt, or the file could not be recovered because the alternate " & _
                   "copy or log was absent or corrupt."
      Case REGERREURREGISTRYIOFAILED
         urerRegErreur.strErreurDescription = "An I/O operation initiated by the Registry failed unrecoverably. " & _
                   "The Registry could not read in, or write out, or flush, one of the files " & _
                   "that contain the system's image of the Registry."
      Case REGERREURNOTREGISTRYFILE
         urerRegErreur.strErreurDescription = "The system has attempted to load or restore a file into the Registry, but the " & _
                   "specified file is not in a Registry file format."
      Case REGERREURKEYDELETED
         urerRegErreur.strErreurDescription = "Illegal operation attempted on a Registry key which has been marked for deletion."
      Case REGERREURNOLOGSPACE
         urerRegErreur.strErreurDescription = "System could not allocate the required space in a Registry log."
      Case REGERREURKEYHASCHILDREN
         urerRegErreur.strErreurDescription = "Cannot create a symbolic link in a Registry key that already " & _
                   "has subkeys or values."
      Case REGERREURCHILDMUSTBEVOLATILE
         urerRegErreur.strErreurDescription = "Cannot create a stable subkey under a volatile parent key."
      Case REGERREURRXACTINVALIDSTATE
         urerRegErreur.strErreurDescription = "The transaction state of a Registry subtree is incompatible with the " & _
                   "requested operation."
   End Select
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, "ACCESS BASE DE REGISTRE", "ERREUR", vbNullString, urerRegErreur
   ' Fin
AjouteRegErreur_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
AjouteRegErreur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEPROCEDURE, "AjouteRegErreur", vbNullString, Err
   ' Je Continue
   Resume AjouteRegErreur_Exit
   ' Fin
End Sub

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
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEPROCEDURE, "Class_Initialize", vbNullString, Err
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
   gfloLogDeltaApi.AjouteErreur App, LOGFEUILLETYPE, LOGFEUILLENOM, LIBELLEPROCEDURE, "Class_Terminate", vbNullString, Err
   ' Je Continue
   Resume Class_Terminate_Exit
   ' Fin
End Sub

