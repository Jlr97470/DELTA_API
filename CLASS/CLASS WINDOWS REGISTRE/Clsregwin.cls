VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegWin"
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
'***        clsRegWin                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      ----------    Pour La Gestion De La Base De Registre Windows  --------   ***
'***            1 - La declaration de constantes , de function et procedure de    ***
'***                 l'api externe propre a cette class                                         ***
'***            2 - Enumere,Ouvre,Ferme,Supprime,Ecriture de Cles                  ***
'***                 De La Base De Registre                                                       ***
'***            2 - Enumere,Lecture,Supprime,Ecriture de Valeurs                     ***
'***                 De La Base De Registre                                                       ***
'***            3 - Convertir Les Valeurs De Registre En Valeurs Visual Basic   ***
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
Private Const mconFEUILLETYPE = FEUILLECLASS                                                     ' La feuille du source
Private Const mconFEUILLENOM = "clsRegWin"                                                 ' Je me rapelle le nom de la Feuille

'******************************************************************************
'***    Constante Qui Defini Si Les Cles Sont Volatiles                                 ***
'******************************************************************************
Private Enum REGOPTIONCONSTANTES
   REG_OPTION_NON_VOLATILE = 0
   REG_OPTION_VOLATILE = 1
   REG_OPTION_BACKUP_RESTORE = 2
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Controls D'Access A La Base De Registre   ***
'******************************************************************************
Private Enum REGCONTROLCONSTANTES
   READ_CONTROL = &H20000
   Synchronize = &H100000
   STANDARD_RIGHTS_ALL = &H1F0000
   STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
   STANDARD_RIGHTS_READ = (READ_CONTROL)
   STANDARD_RIGHTS_REQUIRED = &HF0000
   STANDARD_RIGHTS_WRITE = (READ_CONTROL)
   KEY_CREATE_LINK = &H20
   KEY_CREATE_SUB_KEY = &H4
   KEY_ENUMERATE_SUB_KEYS = &H8
   KEY_NOTIFY = &H10
   KEY_QUERY_VALUE = &H1
   KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not Synchronize))
   KEY_EXECUTE = (KEY_READ)
   KEY_SET_VALUE = &H2
   KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not Synchronize))
   KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not Synchronize))
 End Enum

'******************************************************************************
'***    Constante Qui Defini Le Resultat De La Fonction RegEcritureCle        ***
'******************************************************************************
Private Enum REGDISPOSITIONCONSTANTES
   REG_CREATED_NEW_KEY = 0                                                                      ' Valeur a definir
   REG_OPENED_EXISTING_KEY = 1
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Erreur D'Access A La Base De Registre      ***
'******************************************************************************
Private Enum REGERREURCONSTANTES
   ERROR_SUCCESS = 0&
   ERROR_INVALID_KEY = 2&
   ERROR_ACCESS_DENIED = 5&
   ERROR_OUTOFMEMORY = 14&
   ERROR_INVALID_PARAMETER = 87&
   ERROR_MORE_DATA = 234&
   ERROR_NO_MORE_ITEMS = 259&
   ERROR_BADDB = 1009&
   ERROR_BADKEY = 1010&
   ERROR_CANTOPEN = 1011&
   ERROR_CANTREAD = 1012&
   ERROR_CANTWRITE = 1013&
   ERROR_REGISTRY_RECOVERED = 1014&
   ERROR_REGISTRY_CORRUPT = 1015&
   ERROR_REGISTRY_IO_FAILED = 1016&
   ERROR_NOT_REGISTRY_FILE = 1017&
   ERROR_KEY_DELETED = 1018&
   ERROR_NO_LOG_SPACE = 1019&
   ERROR_KEY_HAS_CHILDREN = 1020&
   ERROR_CHILD_MUST_BE_VOLATILE = 1021&
   ERROR_RXACT_INVALID_STATE = 1369&
End Enum

'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Cles sectionss De La Base De Registre         ***
'******************************************************************************
Public Enum REGSECTIONCONSTANTES
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_CURRENT_USER = &H80000001
   HKEY_DYN_DATA = &H80000006
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_USERS = &H80000003
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Types De Valeur De La Base De Registre    ***
'******************************************************************************
Public Enum REGTYPECONSTANTES
   REG_NONE = 0                                                                                             ' No value type
   REG_SZ = 1                                                                                                  ' Unicode nul terminated string
   REG_EXPAND_SZ = 2                                                                                   ' Unicode nul terminated string
   REG_BINARY = 3                                                                                         ' Free form binary
   REG_DWORD_LITTLE_ENDIAN = 4                                                                ' 32-bit number (same as REG_DWORD)
   REG_DWORD = 4                                                                                         ' 32-bit number
   REG_DWORD_BIG_ENDIAN = 5                                                                     ' 32-bit number
   REG_LINK = 6                                                                                              ' Symbolic Link (unicode)
   REG_MULTI_SZ = 7                                                                                      ' Multiple Unicode strings
 End Enum

'******************************************************************************
'***    Declaration De Type Private                                                              ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGCONTROLACCESS
    bytAccessRevision As Byte
    bytAccessBz1 As Byte
    intAccessTaille As Integer
    intAccessNombre As Integer
    intAccessBz2 As Integer
End Type

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGSECURITERATTRIBUE
    lngAttribueLongueur As Long
    lngAttribueDescription As Long
    lngAttribueEntete As Long
End Type

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Private Type REGSECURITERDESCRIPTION
    bytDescriptionRevision As Byte
    bytDescriptionBz1 As Byte
    lngDescriptionControl As Long
    lngDescriptionOwner As Long
    lngDescriptionGroup As Long
    urcaDescriptionS As REGCONTROLACCESS
    urcaDescriptionD As REGCONTROLACCESS
End Type

'******************************************************************************
'***    Declaration De Type Priver                                                               ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion  D'Une Erreur D'Access A La Base De Registre ***
'******************************************************************************
Public Type REGERREUR
    lngErreurNumber As Long
    strErreurDescription As String
End Type

'******************************************************************************
'***    Declaration De Type Public                                                               ***
'******************************************************************************

'******************************************************************************
'***    Type Pour La Gestion Du Dernier Access A La Base De Registre        ***
'******************************************************************************
Public Type REGDATE
    lngDateBas As Long
    lngDateHaut As Long
End Type

'******************************************************************************
'***    Type Pour La Gestion  Des Cles Primaires De La Base De Registre    ***
'******************************************************************************
Public Type REGCLEPRIM
   lngPrimSection As REGSECTIONCONSTANTES
   strPrimCle As String
End Type

'******************************************************************************
'***    Type Pour La Gestion                                                                        ***
'******************************************************************************
Public Type REGCLE
   strCleIndex As String
   lngCleEntete As Long
   urdtCleDate As REGDATE
   urcpClePrim As REGCLEPRIM
   blnCleSousEnumerer As Boolean
   colCleSous As Collection
   blnCleValeurEnumerer As Boolean
   colCleValeur As Collection
End Type

'******************************************************************************
'***    Type Pour La Gestion  D'Une Valeur De La Base De Registre              ***
'******************************************************************************
Public Type REGVALEUR
    strValeurIndex As String
    strValeurNom As String
    lngValeurType As REGTYPECONSTANTES
    vntValeurVba As Variant
End Type

'******************************************************************************
'***    Declaration De Fonction Private Externe                                            ***
'******************************************************************************
    
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegOuvreCle(ByVal lngRegSection As REGSECTIONCONSTANTES, ***
'***      ByVal strRegCle As String,                                                             ***
'***      ByVal lngRegOption As REGOPTIONCONSTANTES,                         ***
'***      ByVal lngRegControl As REGCONTROLCONSTANTES,                     ***
'***      ByRef lngRegCleEntete As Long)  As REGERREURCONSTANTES     ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ouvrire une cle dans la base de registre                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                                  - La cle SECTION                    ***
'***      strRegCle                                       - La sous cle                            ***
'***      lngRegOption                                  - L'option (Voir ref NT)              ***
'***      lngRegControl                                - Le mode de control d'access   ***
'***      lngRegCleEntete                             - L'entete de cle ouvert             ***
'***    SORTIE:                                                                                           ***
'***      RegOuvreCle                                   - Recupere le status de            ***
'***                                                               l'Operation  externe               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegOuvreCle(lngRegSection, strRegCle, 0,                                      ***
'***     KEY_ALL_ACCESS, lngRegCleEntete)                                              ***
'******************************************************************************
Private Declare Function RegOuvreCle Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String, ByVal lngRegOption As REGOPTIONCONSTANTES, ByVal lngRegControl As REGCONTROLCONSTANTES, ByRef lngRegCleEntete As Long) As REGERREURCONSTANTES

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegFermeCle(ByVal lngRegCleEntete As Long)                                ***
'***      As REGERREURCONSTANTES                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour fermer une cle de la base de registre                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                              - L'entete de cle ouvert            ***
'***    SORTIE:                                                                                           ***
'***      RegFermeCle                                   - Recupere le status de            ***
'***                                                               l'operation externe                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegFermeCle(lngRegCleEntete)                                                       ***
'******************************************************************************
Private Declare Function RegFermeCle Lib "advapi32.dll" Alias "RegCloseKey" (ByVal lngRegCleEntete As Long) As REGERREURCONSTANTES

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegEcritureCle                                                                               ***
'***     (ByVal lngRegSection As REGSECTIONCONSTANTES,                      ***
'***      ByVal strRegCle As String, ByVal lngRegReserver As Long,            ***
'***      ByVal strRegClass As String,                                                         ***
'***      ByVal lngRegOption As REGOPTIONCONSTANTES,                         ***
'***      ByVal lngRegControl As REGCONTROLCONSTANTES,                    ***
'***      ByVal lngRegSecuriter As long, ByRef lngRegCleEntete As Long,    ***
'***      ByRef lngRegDisposition As Long) As REGERREURCONSTANTES   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire une cle dans la base de registre                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                             - La  SECTION                             ***
'***      strRegCle                                     - La  cle                                     ***
'***      lngRegReserver                                                                              ***
'***      strRegClass                                                                                    ***
'***      lngRegOption                                  - L'option (Voir ref NT)               ***
'***      lngRegControl                                - Le mode de control d'access    ***
'***      lngRegSecuriter                                                                              ***
'***      lngRegCleEntete                             - La l'entete de cle ouvert          ***
'***      lngRegDisposition                                                                           ***
'***    SORTIE:                                                                                           ***
'***      RegEcritureCle                                - Recupere le status  de           ***
'***                                                               l'Operation  externe               ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegEcritureCle(lngRegSection, strRegCle, 0, vbNullString, 0,           ***
'***     KEY_ALL_ACCESS, Null, lngRegCleEntete, 0)                                  ***
'******************************************************************************
Private Declare Function RegEcritureCle Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String, ByVal lngRegReserver As Long, ByVal strRegClass As String, ByVal lngRegOption As REGOPTIONCONSTANTES, ByVal lngRegControl As REGCONTROLCONSTANTES, ByVal lngRegSecuriter As Long, ByRef lngRegCleEntete As Long, ByRef lngRegDisposition As Long) As REGERREURCONSTANTES

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegSupprimeCle                                                                             ***
'***      (ByVal lngRegSection As REGSECTIONCONSTANTES,                     ***
'***      ByVal strRegCle As String) As REGERREURCONSTANTES              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour supprimer une liste de cle de la base de registre                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegSection                           - La cle SECTION                           ***
'***      strRegCle                                     - La sous cle                              ***
'***    SORTIE:                                                                                           ***
'***      RegSupprimeCle                            - Recupere le status  de             ***
'***                                                             l'Operation  externe                 ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegSupprimeCle(lngRegSection, strRegCle)                                     ***
'******************************************************************************
Private Declare Function RegSupprimeCle Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngRegSection As REGSECTIONCONSTANTES, ByVal strRegCle As String) As REGERREURCONSTANTES

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegLectureValeur(ByVal lngRegCleEntete As Long,                         ***
'***      ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ***
'***      ByRef lngRegValeurType As REGTYPECONSTANTES,                      ***
'***      ByRef anyNull As Any, ByReflngRegValeurLongueur As Long)          ***
'***      As REGERREURCONSTANTES                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire une valeur entier long de cle de la base de registre            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      lngRegCleEntete                           - La l'entete de cle ouvert           ***
'***      strRegValeurNom                          - Le nom de la valeur                  ***
'***      lngRegReserver                                                                              ***
'***      lngRegValeurType                        - Type de valeur                          ***
'***      anyNull                                         - Valeur Null                               ***
'***      lngRegValeurLongueur                  - Longueur de la valeur                ***
'***                                                           de la  Cle                                  ***
'***    SORTIE:                                                                                           ***
'***      RegLectureValeur                        - Recupere le status de                ***
'***                                                           l'Operation                                ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***     RegLectureValeur(lngRegCleEntete, strRegValeurNom, 0,                ***
'***     lngValeurType, vntValeur, lngRegValeurLongueur)                           ***
'***     lngRegValeurType, Null, lngRegValeurLongueur)                              ***
'******************************************************************************
Private Declare Function RegLectureValeur Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByRef lngRegValeurType As REGTYPECONSTANTES, ByRef anyNull As Any, ByRef lngRegValeurLongueur As Long) As REGERREURCONSTANTES

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      RegLectureValeurChaine(ByVal lngRegCleEntete As Long,               ***
'***      ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ***
'***      ByRef lngRegValeurType As REGTYPECONSTANTES,                     ***
'***      ByVal strRegValeur As String,                                                        ***
'***      ByRef lngRegValeurLongueur As Long)                                            ***
'***      As REGERREURCONSTANTES                                                        ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour lire une valeur chaine de cle de la base de registre                  ***
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
Private Declare Function RegLectureValeurChaine Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByRef lngRegValeurType As REGTYPECONSTANTES, ByVal strRegValeur As String, ByRef lngRegValeurLongueur As Long) As REGERREURCONSTANTES

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
Private Declare Function RegEcritureValeur Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByVal lngRegValeurType As REGTYPECONSTANTES, ByVal lngRegValeur As Long, ByVal lngRegValeurLongueur As Long) As REGERREURCONSTANTES

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
Private Declare Function RegEcritureValeurChaine Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String, ByVal lngRegReserver As Long, ByVal lngRegValeurType As REGTYPECONSTANTES, ByVal strRegValeur As String, ByVal lngRegValeurLongueur As Long) As REGERREURCONSTANTES

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
Private Declare Function RegSupprimeValeur Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngRegCleEntete As Long, ByVal strRegValeurNom As String) As REGERREURCONSTANTES

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
Private Declare Function RegEnumereCle Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal lngRegCleEntete As Long, ByVal lngRegCleIndex As Long, ByVal strRegCle As String, ByRef lngRegCleLongueur As Long, ByVal lngRegReserver As Long, ByVal lngRegClass As Long, ByRef lngRegClassLongueur As Long, ByRef urdhRegCleDateHeure As REGDATE) As REGERREURCONSTANTES

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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CRegValeurVbaValeur", _
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, _
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurbytVbaValeur", strRegValeur, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurintVbaValeur", strRegValeur, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurlngVbaValeur", strRegValeur, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurstrVbaValeurTab", strRegValeur, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrRegValeurbytVbaValeurTab", strRegValeur, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "ClngVbaValeurstrRegValeur", CStr(lngVbaValeur), Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CstrVbaValeurTabstrRegValeur", strVbaValeurTableau(0), Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "CbytVbaValeurTabstrRegValeur", CStr(bytVbaValeurTableau(0)), Err
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
         lngRegWinOperation = RegOuvreCle(urclRegCle.urcpClePrim.lngPrimSection, urclRegCle.urcpClePrim.strPrimCle, 0, KEY_ALL_ACCESS, urclRegCle.lngCleEntete)
         ' Je regarde le statut de l'operation
         Select Case lngRegWinOperation
            Case ERROR_SUCCESS
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
                   Case ERROR_INVALID_KEY
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "OuvreCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
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
                Case ERROR_SUCCESS
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
                       Case ERROR_INVALID_KEY
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "FermeCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
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
                        Case ERROR_SUCCESS
                            ' J'ai enumerer une sous cle
                            ' J'ai  enumerer les cles
                            EnumereCle = True
                        Case ERROR_NO_MORE_ITEMS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EnumereCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
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
Attribute EcritureCle.VB_Description = "Pour ecrire une cle dans la base de registre"
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
         lngRegWinOperation = RegEcritureCle(urclRegCle.urcpClePrim.lngPrimSection, urclRegCle.urcpClePrim.strPrimCle, vbEmpty, REG_NONE, vbEmpty, KEY_ALL_ACCESS, vbEmpty, urclRegCle.lngCleEntete, vbEmpty)
         ' Je regarde le statut de l'operation
         Select Case lngRegWinOperation
            Case ERROR_SUCCESS
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
                   Case ERROR_INVALID_KEY
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EcritureCle", CStr(lngRegSection) & "," & CStr(strRegCle), Err
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
Attribute SupprimeCle.VB_Description = "Pour supprimer une cle de la base de registre"
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
                           Case ERROR_SUCCESS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "SupprimeCle", CStr(lngRegSection) & "," & strRegCle, Err
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
Attribute LectureValeur.VB_Description = "Pour lire une valeur dans la base de registre"
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
                       Case ERROR_SUCCESS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "LectureValeur", CStr(lngRegSection) & "," & CStr(strRegCle) & "," & CStr(strRegValeurNom), Err
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
Attribute EcritureValeur.VB_Description = "Pour ecrire une valeur dans la base de registre"
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
            Case ERROR_SUCCESS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EcritureValeur", CStr(lngRegSection) & "," & CStr(strRegCle) & "," & CStr(strRegValeurNom) & "," & CStr(vntVbaValeur), Err
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
Attribute SupprimeValeur.VB_Description = "Pour supprimer une valeur de la base de registre"
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
                        Case ERROR_SUCCESS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "SupprimeValeur", CStr(lngRegSection) & "," & strRegCle & "," & strRegValeurNom, Err
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
Attribute EnumereValeur.VB_Description = "Pour enumerer des valeurs de la base de registre"
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
                            Case ERROR_SUCCESS
                                ' J'ai enumerer une sous cle
                                ' J'ai  enumerer les cles
                                EnumereValeur = True
                            Case ERROR_NO_MORE_ITEMS
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
                                   Case ERROR_SUCCESS
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EnumereValeur", CStr(lngRegSection) & "," & strRegCle, Err
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
      Case ERROR_ACCESS_DENIED
         urerRegErreur.strErreurDescription = "Access Denier"
      Case ERROR_OUTOFMEMORY
         urerRegErreur.strErreurDescription = "Pas assez de memoire"
      Case ERROR_INVALID_PARAMETER
         urerRegErreur.strErreurDescription = "Parametre invalide"
      Case ERROR_MORE_DATA
         urerRegErreur.strErreurDescription = "Plus de donner"
      Case ERROR_NO_MORE_ITEMS
         urerRegErreur.strErreurDescription = "Plus d'elements"
      Case ERROR_BADDB
         urerRegErreur.strErreurDescription = "The configuration registry database is corrupt."
      Case ERROR_BADKEY
         urerRegErreur.strErreurDescription = "The configuration registry key is invalid."
      Case ERROR_CANTOPEN
         urerRegErreur.strErreurDescription = "The configuration registry key could not be opened."
      Case ERROR_CANTREAD
         urerRegErreur.strErreurDescription = "The configuration registry key could not be read."
      Case ERROR_CANTWRITE
         urerRegErreur.strErreurDescription = "The configuration registry key could not be written."
      Case ERROR_REGISTRY_RECOVERED
         urerRegErreur.strErreurDescription = "One of the files in the Registry database had to be recovered " & _
                   "by use of a log or alternate copy. The recovery was successful."
      Case ERROR_REGISTRY_CORRUPT
         urerRegErreur.strErreurDescription = "The Registry is corrupt. The structure of one of the files that contains " & _
                   "Registry data is corrupt, or the system's image of the file in memory " & _
                   "is corrupt, or the file could not be recovered because the alternate " & _
                   "copy or log was absent or corrupt."
      Case ERROR_REGISTRY_IO_FAILED
         urerRegErreur.strErreurDescription = "An I/O operation initiated by the Registry failed unrecoverably. " & _
                   "The Registry could not read in, or write out, or flush, one of the files " & _
                   "that contain the system's image of the Registry."
      Case ERROR_NOT_REGISTRY_FILE
         urerRegErreur.strErreurDescription = "The system has attempted to load or restore a file into the Registry, but the " & _
                   "specified file is not in a Registry file format."
      Case ERROR_KEY_DELETED
         urerRegErreur.strErreurDescription = "Illegal operation attempted on a Registry key which has been marked for deletion."
      Case ERROR_NO_LOG_SPACE
         urerRegErreur.strErreurDescription = "System could not allocate the required space in a Registry log."
      Case ERROR_KEY_HAS_CHILDREN
         urerRegErreur.strErreurDescription = "Cannot create a symbolic link in a Registry key that already " & _
                   "has subkeys or values."
      Case ERROR_CHILD_MUST_BE_VOLATILE
         urerRegErreur.strErreurDescription = "Cannot create a stable subkey under a volatile parent key."
      Case ERROR_RXACT_INVALID_STATE
         urerRegErreur.strErreurDescription = "The transaction state of a Registry subtree is incompatible with the " & _
                   "requested operation."
   End Select
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, "ACCESS BASE DE REGISTRE", "ERREUR", vbNullString, urerRegErreur
   ' Fin
AjouteRegErreur_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
AjouteRegErreur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONPROCEDURE, "AjouteRegErreur", vbNullString, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONPROCEDURE, "Class_Initialize", vbNullString, Err
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
   gfloLogDeltaApi.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONPROCEDURE, "Class_Terminate", vbNullString, Err
   ' Je Continue
   Resume Class_Terminate_Exit
   ' Fin
End Sub

