VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsPanConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'
'Public Sub FormatFloppy()
'
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,SHFormatDrive"
'
' Call CenterDialog("Format - 3½ Floppy (A:)")
' x = Shell(Procs, 1)
' Call CenterDialog("Format - 3½ Floppy (A:)")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Sounds()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl " & """Sounds"""
'
' Call CenterDialog("Sounds Properties")
' x = Shell(Procs, 1)
' Call CenterDialog("Sounds Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub AddNewHardware()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl " & """Add New Hardware"""""
'
'CenterDialog ("Add New Hardware Wizard")
' x = Shell(Procs, 1)
'CenterDialog ("Add New Hardware Wizard")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub AddNewSoftWare()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl"
'
'CenterDialog ("Add/Remove Programs Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Add/Remove Programs Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub DateTime()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL Timedate.cpl"
'
'CenterDialog ("Date/Time Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Date/Time Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub DesktopThemes()
'On Error GoTo notinstalled:
'
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL Themes.cpl"
'
'CenterDialog ("Desktop Themes")
' x = Shell(Procs, 1)
'CenterDialog ("Desktop Themes")
'
'k = LockWindowUpdate(0)
'notinstalled:
'k = LockWindowUpdate(0)
'MsgBox "Microsoft Plus not installed"
'End Sub
'
'Public Sub Display()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL Desk.cpl"
'
'CenterDialog ("Display Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Display Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Internet()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
'
'CenterDialog ("Internet Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Internet Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Keyboard()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL main.cpl " & """KeyBoard"""""
'
'Call CenterDialog("Keyboard Properties")
' x = Shell(Procs, 1)
'Call CenterDialog("Keyboard Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Exchange()
'On Error GoTo notinstalled:
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL mlcfg32.cpl"
'
'CenterDialog ("MS Exchange Settings Properties")
' x = Shell(Procs, 1)
'CenterDialog ("MS Exchange Settings Properties")
'
'k = LockWindowUpdate(0)
'notinstalled:
'k = LockWindowUpdate(0)
'MsgBox "Exchange not installed."
'
'End Sub
'
'Public Sub Modems()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL modem.cpl"
'
'CenterDialog ("Modems Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Modems Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Mouse()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL main.cpl"
'
'CenterDialog ("Mouse Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Mouse Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Multimedia()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl"
'
'CenterDialog ("Multimedia Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Multimedia Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Network()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl"
'
'CenterDialog ("NetWork")
' x = Shell(Procs, 1)
'CenterDialog ("NetWork")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub PassWords()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL password.cpl"
'
'CenterDialog ("Passwords Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Passwords Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub Regional()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL intl.cpl"
'
'CenterDialog ("Regional Settings Properties")
' x = Shell(Procs, 1)
'CenterDialog ("Regional Settings Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub System()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl"
'
'CenterDialog ("System Properties")
' x = Shell(Procs, 1)
'CenterDialog ("System Properties")
'
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub TweakUI()
'On Error GoTo notinstalled:
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe shell32.dll,Control_RunDLL tweakui.cpl"
'
'CenterDialog ("Tweak UI")
' x = Shell(Procs, 1)
'CenterDialog ("Tweak UI")
'
'k = LockWindowUpdate(0)
'notinstalled:
'MsgBox "TweakUI not installed"
'k = LockWindowUpdate(0)
'End Sub
'
'Public Sub DiskCopy1()
'Dim sBuffer As String, Windir As String, Procs As String, x
'Dim lResult As Long
'sBuffer = String$(255, 0)
'lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
'
'Windir = Trim(sBuffer)
'
'Procs = Left(Windir, lResult) & "\rundll32.exe Diskcopy.dll,DiskCopyRunDll"
'
' Call CenterDialog("Copy Disk")
' x = Shell(Procs, 1)
'Call CenterDialog("Copy Disk")
'
'k = LockWindowUpdate(0)
'End Sub
'
