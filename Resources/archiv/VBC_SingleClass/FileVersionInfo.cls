VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileVersionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public NotInheritable Class FileVersionInfo
'          Inherits System.Object
'     Member von: System.Diagnostics
'Zusammenfassung:
' Stellt Versionsinformationen zu einer physischen
' Datei auf einem Datentr�ger bereit.
#If defUnicode Then
  Private Declare Function pGetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
  Private Declare Function pGetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
  Private Declare Function pVerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
  Private Declare Function pVerLanguageName Lib "kernel32.dll" Alias "VerLanguageNameW" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
  Private Declare Function plstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function plstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal dst As Long, ByVal src As Long) As Long
#Else
  Private Declare Function pGetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
  Private Declare Function pGetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
  Private Declare Function pVerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
  Private Declare Function pVerLanguageName Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
  Private Declare Function plstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
  Private Declare Function plstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal dst As Long, ByVal src As Long) As Long
#End If
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef pDst As Any, ByRef pSrc As Long, ByVal Length As Long)

Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&

Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000
Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA

Private Const VFT2_FONT_RASTER = &H1
Private Const VFT2_FONT_VECTOR = &H2
Private Const VFT2_FONT_TRUETYPE = &H3

Private Const MAX_PATH = 260

'the prefix dw means doubleword
'the versions here are only words so drop the "d"
Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  wStrucVersionl     As Integer ' minor
  wStrucVersionh     As Integer 'major
  wFileVersionMSl    As Integer ' minor
  wFileVersionMSh    As Integer 'major
  wFileVersionLSl    As Integer '   private
  wFileVersionLSh    As Integer '  build
  wProductVersionMSl As Integer ' minor
  wProductVersionMSh As Integer 'major
  wProductVersionLSl As Integer '   private
  wProductVersionLSh As Integer '  build
  dwFileFlagsMask    As Long
  dwFileFlags        As Long
  dwFileOS           As Long
  dwFileType         As Long
  dwFileSubtype      As Long
  dwFileDateMS       As Long
  dwFileDateLS       As Long
End Type
Private mComments         As String
Private mCompanyName      As String
Private mFileDescription  As String
Private mVSFileInfo       As VS_FIXEDFILEINFO
Private mFileName         As String
Private mFileVersion      As String
Private mInternalName     As String
Private mLanguage         As String
Private mLegalCopyright   As String
Private mLegalTrademarks  As String
Private mOriginalFilename As String
Private mPrivateBuild     As String
Private mProductName      As String
Private mProductVersion   As String
Private mSpecialBuild     As String

'der Konstruktor nur Private, warum eigentlich? das wei� kein Mensch
Friend Sub NewC(StrFileName As String)
  mFileName = StrFileName
End Sub
'in ein Modul ModConstructors kopieren
Friend Function New_FileVersionInfo(StrFileName As String) As FileVersionInfo
  Set New_FileVersionInfo = New FileVersionInfo
  Call New_FileVersionInfo.NewC(StrFileName)
End Function

'Public Shared Function GetVersionInfo(ByVal fileName As String) As System.Diagnostics.FileVersionInfo
Public Function GetVersionInfo(ByVal FileName As String) As FileVersionInfo
  If (LenB(Dir$(FileName)) = 0) Then
    MsgBox "FileNotFoundException: " & FileName
    Exit Function
  End If
  Dim num2 As Long: num2 = pGetFileVersionInfoSize(FileName, 0)
  Dim info1 As FileVersionInfo: Set info1 = New_FileVersionInfo(FileName)
  If (num2 = 0) Then
    Set GetVersionInfo = info1: Exit Function
  End If
  Dim buffer1() As Byte: ReDim buffer1(0 To num2 - 1)
  Dim numRef1 As Long: numRef1 = VarPtr(buffer1(0))
    Dim ptr1 As Long: ptr1 = numRef1 ': Set ptr1 = New_IntPtr(numRef1)
    If CBool(pGetFileVersionInfo(FileName, 0, num2, ByVal ptr1)) Then
      Debug.Print buffer1
      Dim num3 As Long: num3 = GetVarEntry(ptr1)
      If Not info1.GetVersionInfoForCodePage(ptr1, ConvertTo8DigitHex(num3)) Then
        'Wenn die Sprache nicht geklappt hat, dann noch mit den drei anderen probieren
        Dim num4, numArray1: numArray1 = Array(&H40904B0, &H40904E4, &H4090000)
        For Each num4 In numArray1
          If (num4 <> num3) Then
            If info1.GetVersionInfoForCodePage(ptr1, ConvertTo8DigitHex(num4)) Then
              Exit For
            End If
          End If
        Next
      End If
    End If
  Set GetVersionInfo = info1
End Function

'Private Shared Function ConvertTo8DigitHex(ByVal value As Integer) As String
Friend Function ConvertTo8DigitHex(ByVal value As Long) As String
  Dim StrVal As String: StrVal = Hex$(value)
  ConvertTo8DigitHex = String$(8 - Len(StrVal), "0") & StrVal
End Function

'Private Shared Function GetVarEntry(ByVal memPtr As IntPtr) As Integer
Friend Function GetVarEntry(ByVal memPtr As Long) As Long
  Dim num1 As Long
  Dim ptr1 As Long
  If pVerQueryValue(ByVal memPtr, "\VarFileInfo\Translation", ByVal VarPtr(ptr1), num1) Then
    GetVarEntry = ShL(ReadInt16(ptr1), 16) + CLng(ReadInt16(ptr1, 2))
    Exit Function
  End If
  GetVarEntry = &H40904E4  '67699940
End Function

'Private Shared Function GetFileVersionLanguage(ByVal memPtr As IntPtr) As String
Friend Function GetFileVersionLanguage(ByVal memPtr As Long) As String
  Dim num1 As Long: num1 = ShR(GetVarEntry(memPtr), 16)
  Dim buffer As String * 256
  Dim L As Long
  L = pVerLanguageName(num1, buffer, 256)
  GetFileVersionLanguage = Left$(buffer, L)
End Function

'Private Shared Function GetFileVersionString(ByVal memPtr As IntPtr, ByVal name As String) As String
Friend Function GetFileVersionString(ByVal memPtr As Long, ByVal name As String) As String
  Dim num1 As Long
  Dim ptr1 As Long
  If pVerQueryValue(ByVal memPtr, name, ptr1, num1) Then
    If (ptr1 <> 0) Then
      GetFileVersionString = PtrToString(ptr1)
    End If
  End If
End Function

Friend Function GetVersionInfoForCodePage(ByVal memIntPtr As Long, ByVal codepage As String) As Boolean
  Dim text1 As String: text1 = "\\StringFileInfo\\"
  mCompanyName = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "CompanyName")
  mFileDescription = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "FileDescription")
  mFileVersion = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "FileVersion")
  mInternalName = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "InternalName")
  mLegalCopyright = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "LegalCopyright")
  mOriginalFilename = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "OriginalFilename")
  mProductName = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "ProductName")
  mProductVersion = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "ProductVersion")
  mComments = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "Comments")
  mLegalTrademarks = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "LegalTrademarks")
  mPrivateBuild = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "PrivateBuild")
  mSpecialBuild = GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "SpecialBuild")
  mLanguage = GetFileVersionLanguage(memIntPtr)
  mVSFileInfo = GetFixedFileInfo(memIntPtr)
'  mFileMajor = HIWORD(mVSFileInfo.dwFileVersionMS)
'  mFileMinor = LOWORD(mVSFileInfo.dwFileVersionMS)
'  mFileBuild = HIWORD(mVSFileInfo.dwFileVersionLS)
'  mFilePrivate = LOWORD(mVSFileInfo.dwFileVersionLS)
'  mProductMajor = HIWORD(mVSFileInfo.dwProductVersionMS)
'  mProductMinor = LOWORD(mVSFileInfo.dwProductVersionMS)
'  mProductBuild = HIWORD(mVSFileInfo.dwProductVersionLS)
'  mProductPrivate = LOWORD(mVSFileInfo.dwProductVersionLS)
'  mFileFlags = vs_fixedfileinfo1.dwFileFlags
  'Return (Not Me.fileVersion Is String.Empty)
  GetVersionInfoForCodePage = Not (Len(mFileVersion) = 0)
End Function

'Private Shared Function GetFixedFileInfo(ByVal memPtr As IntPtr) As VS_FIXEDFILEINFO
Friend Function GetFixedFileInfo(ByVal memPtr As Long) As VS_FIXEDFILEINFO
  Dim num1 As Long
  Dim ptr As Long
  Call pVerQueryValue(ByVal memPtr, "\", ptr, num1)
  Call PtrToStructure(ptr, VarPtr(GetFixedFileInfo), LenB(GetFixedFileInfo))
End Function

'##############################'    Math    '##############################'
Private Function ShL(n As Integer, c As Long) As Long
  ShL = n * 2 ^ c
End Function
Private Function ShR(n As Long, c As Long) As Long
  ShR = n \ 2 ^ c
End Function

'##############################'    Util    '##############################'
'Private Shared Function HIWORD(ByVal dword As Integer) As Integer
Friend Function HIWORD(ByVal n As Long) As Long
  'HIWORD = (Math.ShR(n, 16) And 65535)
  'oder auch:
  HIWORD = (n \ 65536) And &HFFFF&
End Function
'Private Shared Function LOWORD(ByVal dword As Integer) As Integer
Friend Function LOWORD(ByVal n As Long) As Long
  'LOWORD = (n And 65535)
  'oder auch:
  LOWORD = n And &HFFFF&
End Function

'##############################'  Marshal   '##############################'
Public Function ReadInt16(ByVal ptr As Long, Optional ByVal ofs As Long) As Integer
  '                        pDst             , pSrc
  Call RtlMoveMemory(ByVal VarPtr(ReadInt16), ByVal (ptr + ofs), 2)
End Function
Public Function PtrToString(ByVal ptr As Long, Optional ByVal sLen As Long) As String
  If (ptr = 0) Then
    MsgBox "Marshal.PtrToString: Ptr=Nullpointer"
    Exit Function
  End If
  Dim num1 As Long: num1 = plstrlen(ptr)
  PtrToString = Space$(num1)
  Call plstrcpy(StrPtr(PtrToString), ptr)
#If defUnicode Then
  'ist es dann schon der richtige String?
  MsgBox PtrToString
#Else
  PtrToString = Left$(StrConv(PtrToString, vbUnicode), num1)
#End If
End Function
'Public Shared Function PtrToStructure(ByVal ptr As System.IntPtr, ByVal structureType As System.Type) As Object
Public Sub PtrToStructure(ByVal ptr As Long, ByVal pStruct As Long, ByVal LenBStruct As Long)
  'hier wird nicht der Pointer �bertragen, sondern vielmehr der gesamte Speicherbereich
  'in das Objekt structure hineinkopiert
  Call RtlMoveMemory(ByVal pStruct, ByVal ptr, LenBStruct)
End Sub


'##############################'   My Properties   '##############################'
'Alle Properties sind ReadOnly
Public Property Get FileName() As String: FileName = mFileName: End Property
Public Property Get InternalName() As String:  InternalName = mInternalName: End Property
Public Property Get OriginalFilename() As String:   OriginalFilename = mOriginalFilename: End Property
Public Property Get Comments() As String:   Comments = mComments: End Property
Public Property Get CompanyName() As String:  CompanyName = mCompanyName: End Property
Public Property Get Language() As String:  Language = mLanguage: End Property
Public Property Get LegalCopyright() As String:  LegalCopyright = mLegalCopyright: End Property
Public Property Get LegalTrademarks() As String:  LegalTrademarks = mLegalTrademarks: End Property
Public Property Get PrivateBuild() As String:  PrivateBuild = mPrivateBuild: End Property
Public Property Get SpecialBuild() As String:  SpecialBuild = mSpecialBuild: End Property

'##############################'    Fileversion    '##############################'
Public Property Get FileVersion() As String:   FileVersion = mFileVersion: End Property
Public Property Get FileMajorPart() As Long:   FileMajorPart = CLng(mVSFileInfo.wFileVersionMSh): End Property
Public Property Get FileMinorPart() As Long:   FileMinorPart = CLng(mVSFileInfo.wFileVersionMSl): End Property
Public Property Get FileBuildPart() As Long:   FileBuildPart = CLng(mVSFileInfo.wFileVersionLSh): End Property
Public Property Get FilePrivatePart() As Long: FilePrivatePart = CLng(mVSFileInfo.wFileVersionLSl): End Property

'------------------------------------------------------------------
Public Property Get FileDescription() As String:  FileDescription = mFileDescription: End Property
Public Property Get ProductName() As String:  ProductName = mProductName: End Property
'------------------------------------------------------------------

'##############################'  Productversion   '##############################'
Public Property Get ProductVersion() As String:   ProductVersion = mProductVersion: End Property
Public Property Get ProductMajorPart() As Long:   ProductMajorPart = CLng(mVSFileInfo.wProductVersionMSh): End Property
Public Property Get ProductMinorPart() As Long:   ProductMinorPart = CLng(mVSFileInfo.wProductVersionMSl): End Property
Public Property Get ProductBuildPart() As Long:   ProductBuildPart = CLng(mVSFileInfo.wProductVersionLSh): End Property
Public Property Get ProductPrivatePart() As Long: ProductPrivatePart = CLng(mVSFileInfo.wProductVersionLSl): End Property


'##############################'    Bool Props     '##############################'
'Public ReadOnly Property IsDebug() As Boolean
Public Property Get IsDebug() As Boolean
  IsDebug = ((mVSFileInfo.dwFileFlags And VS_FF_DEBUG) <> 0)
End Property
'Public ReadOnly Property IsPreRelease() As Boolean
Public Property Get IsPreRelease() As Boolean
  IsPreRelease = ((mVSFileInfo.dwFileFlags And VS_FF_PRERELEASE) <> 0)
End Property
'Public ReadOnly Property IsPatched() As Boolean
Public Property Get IsPatched() As Boolean
  IsPatched = ((mVSFileInfo.dwFileFlags And VS_FF_PATCHED) <> 0)
End Property
'Public ReadOnly Property IsPrivateBuild() As Boolean
Public Property Get IsPrivateBuild() As Boolean
  IsPrivateBuild = ((mVSFileInfo.dwFileFlags And VS_FF_PRIVATEBUILD) <> 0)
End Property
'noch ein zus�tzliches Property
Public Property Get IsInfoInferred() As Boolean
  IsInfoInferred = ((mVSFileInfo.dwFileFlags And VS_FF_INFOINFERRED) <> 0)
End Property
'Public ReadOnly Property IsSpecialBuild() As Boolean
Public Property Get IsSpecialBuild() As Boolean
  IsSpecialBuild = ((mVSFileInfo.dwFileFlags And VS_FF_SPECIALBUILD) <> 0)
End Property

'More Additional Properties
Public Property Get FileOS() As String
  Select Case mVSFileInfo.dwFileOS
    Case VOS_DOS_WINDOWS16: FileOS = "DOS-Win16"
    Case VOS_DOS_WINDOWS32: FileOS = "DOS-Win32"
    Case VOS_OS216_PM16:    FileOS = "OS/2-16 PM-16"
    Case VOS_OS232_PM32:    FileOS = "OS/2-16 PM-32"
    Case VOS_NT_WINDOWS32:  FileOS = "Win32-NT"
    Case Else:              FileOS = "Unbekannt"
  End Select
End Property

Public Property Get FileType() As String
Dim STyp As String
  Select Case mVSFileInfo.dwFileType
    Case VFT_APP:                FileType = "Application"
    Case VFT_DLL:                FileType = "Dynamic Link Library"
    Case VFT_DRV:                FileType = "Driver"
      Select Case mVSFileInfo.dwFileSubtype
        Case VFT2_DRV_PRINTER:     STyp = "Printer drv"
        Case VFT2_DRV_KEYBOARD:    STyp = "Keyboard drv"
        Case VFT2_DRV_LANGUAGE:    STyp = "Language drv"
        Case VFT2_DRV_DISPLAY:     STyp = "Display drv"
        Case VFT2_DRV_MOUSE:       STyp = "Mouse drv"
        Case VFT2_DRV_NETWORK:     STyp = "Network drv"
        Case VFT2_DRV_SYSTEM:      STyp = "System drv"
        Case VFT2_DRV_INSTALLABLE: STyp = "Installable"
        Case VFT2_DRV_SOUND:       STyp = "Sound drv"
        Case VFT2_DRV_COMM:        STyp = "Comm drv"
        Case VFT2_UNKNOWN:         STyp = "Unknown"
      End Select
    Case VFT_FONT:               FileType = "Font"
      Select Case mVSFileInfo.dwFileSubtype
        Case VFT2_FONT_RASTER:     STyp = "Raster Font"
        Case VFT2_FONT_VECTOR:     STyp = "Vector Font"
        Case VFT2_FONT_TRUETYPE:   STyp = "TrueType Font"
      End Select
    Case VFT_VXD:                FileType = "VxD"
    Case VFT_STATIC_LIB:         FileType = "Lib"
    Case Else:                   FileType = "Unbekannt"
  End Select
  If Len(STyp) > 0 Then FileType = FileType & ", " & STyp
End Property

'Public Overrides Function ToString() As String
Public Function ToString() As String
'Zusammenfassung:
' Gibt eine unvollst�ndige Liste der Eigenschaften in
' System.Diagnostics.FileVersionInfo und deren Werte zur�ck.
'R�ckgabewerte:
' Eine Liste der folgenden Eigenschaften in dieser Klasse und deren Werte:
' System.Diagnostics.FileVersionInfo.FileName ,
' System.Diagnostics.FileVersionInfo.InternalName ,
' System.Diagnostics.FileVersionInfo.OriginalFilename ,
' System.Diagnostics.FileVersionInfo.FileVersion ,
' System.Diagnostics.FileVersionInfo.FileDescription ,
' System.Diagnostics.FileVersionInfo.ProductName ,
' System.Diagnostics.FileVersionInfo.ProductVersion ,
' System.Diagnostics.FileVersionInfo.IsDebug ,
' System.Diagnostics.FileVersionInfo.IsPatched ,
' System.Diagnostics.FileVersionInfo.IsPreRelease ,
' System.Diagnostics.FileVersionInfo.IsPrivateBuild ,
' System.Diagnostics.FileVersionInfo.IsSpecialBuild .
'Wenn keine Datei mit dem angegebenen Namen gefunden wurde, enth�lt diese
'Liste nur den Namen der angeforderten Datei. Boolesche Werte sind auf false
'und alle anderen Eintr�ge auf null festgelegt.
'
'  Dim SB As StringBuilder: Set SB = New_StringBuilder(, , , 128)
'  Call SB.Append("File:             ").Append(mFileName).Append(vbCrLf)
'  Call SB.Append("InternalName:     ").Append(mInternalName).Append(vbCrLf)
'  Call SB.Append("OriginalFilename: ").Append(mOriginalFilename).Append(vbCrLf)
'  Call SB.Append("FileVersion:      ").Append(mFileVersion).Append(vbCrLf)
'  Call SB.Append("FileDescription:  ").Append(mFileDescription).Append(vbCrLf)
'  Call SB.Append("Product:          ").Append(mProductName).Append(vbCrLf)
'  Call SB.Append("ProductVersion:   ").Append(mProductVersion).Append(vbCrLf)
'  Call SB.Append("Debug:            ").Append(CStr(Me.IsDebug)).Append(vbCrLf)
'  Call SB.Append("Patched:          ").Append(CStr(Me.IsPatched)).Append(vbCrLf)
'  Call SB.Append("PreRelease:       ").Append(CStr(Me.IsPreRelease)).Append(vbCrLf)
'  Call SB.Append("PrivateBuild:     ").Append(CStr(Me.IsPrivateBuild)).Append(vbCrLf)
'  Call SB.Append("SpecialBuild:     ").Append(CStr(Me.IsSpecialBuild)).Append(vbCrLf)
'  Call SB.Append("Language:         ").Append(Me.Language).Append(vbCrLf)
'  ToString = SB.ToString
'
'If you have StringBuilder anyway, and need for speed, ok use the upper part.
'well, but we can make it also a bit more independent from other stuff, as
'speed does'nt matter, so use the part below here:
  ToString = ToString & "File:             " & mFileName & vbCrLf
  ToString = ToString & "InternalName:     " & mInternalName & vbCrLf
  ToString = ToString & "OriginalFilename: " & mOriginalFilename & vbCrLf
  ToString = ToString & "FileVersion:      " & mFileVersion & vbCrLf
  ToString = ToString & "FileDescription:  " & mFileDescription & vbCrLf
  ToString = ToString & "Product:          " & mProductName & vbCrLf
  ToString = ToString & "ProductVersion:   " & mProductVersion & vbCrLf
  ToString = ToString & "Debug:            " & CStr(Me.IsDebug) & vbCrLf
  ToString = ToString & "Patched:          " & CStr(Me.IsPatched) & vbCrLf
  ToString = ToString & "PreRelease:       " & CStr(Me.IsPreRelease) & vbCrLf
  ToString = ToString & "PrivateBuild:     " & CStr(Me.IsPrivateBuild) & vbCrLf
  ToString = ToString & "SpecialBuild:     " & CStr(Me.IsSpecialBuild) & vbCrLf
  ToString = ToString & "Language:         " & Me.Language & vbCrLf
  ToString = ToString & "InfoInferred: " & CStr(Me.IsInfoInferred) & vbCrLf
  ToString = ToString & "FileVInt:     " & Me.FileMajorPart & "." & Me.FileMinorPart & "." & Me.FileBuildPart & "." & Me.FilePrivatePart & vbCrLf
  ToString = ToString & "ProdVInt:     " & Me.ProductMajorPart & "." & Me.ProductMinorPart & "." & Me.ProductBuildPart & "." & Me.ProductPrivatePart & vbCrLf
  ToString = ToString & "FileType:     " & Me.FileType & vbCrLf
  ToString = ToString & "FileOS:       " & Me.FileOS
End Function