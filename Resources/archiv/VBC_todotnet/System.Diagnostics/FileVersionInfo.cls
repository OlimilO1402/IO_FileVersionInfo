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
' Datei auf einem Datenträger bereit.
'Private Declare Function GetFileVersionInfoSizeA Lib "Version.dll" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'Private Declare Function GetFileVersionInfoA Lib "Version.dll" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
'Private Declare Function VerQueryValueA Lib "Version.dll" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
'Private Declare Function VerLanguageNameA Lib "kernel32.dll" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
'Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef dst As Any, ByRef src As Long, ByVal Length As Long)
'Private Declare Function GetSystemDirectoryA Lib "kernel32.dll" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Function WritePrivateProfileStringA Lib "kernel32.dll" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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

'Private Type VS_FIXEDFILEINFO
'  dwSignature As Long
'  dwStrucVersionl As Integer
'  dwStrucVersionh As Integer
'  dwFileVersionMSl As Integer
'  dwFileVersionMSh As Integer
'  dwFileVersionLSl As Integer
'  dwFileVersionLSh As Integer
'  dwProductVersionMSl As Integer
'  dwProductVersionMSh As Integer
'  dwProductVersionLSl As Integer
'  dwProductVersionLSh As Integer
'  dwFileFlagsMask As Long
'  dwFileFlags As Long
'  dwFileOS As Long
'  dwFileType As Long
'  dwFileSubtype As Long
'  dwFileDateMS As Long
'  dwFileDateLS As Long
'End Type

Private Type PVersionInfo
  mMinor   As Integer
  mMajor   As Integer
  mPrivate As Integer
  mBuild   As Integer
End Type
Private Type Int32HiLo
  mHi As Integer
  mLo As Integer
End Type
'Private mStrPFN As String 'der Dateiname
'Private mStrFileVersion11 As String 'Major
'Private mStrFileVersion12 As String 'Minor
'Private mStrFileVersion21 As String 'Build
'Private mStrFileVersion22 As String 'Private
'
'Private mStrProdVersion11 As String
'Private mStrProdVersion12 As String
'Private mStrProdVersion21 As String
'Private mStrProdVersion22 As String
'
'Private mStrStrucVersion1 As String
'Private mStrStrucVersion2 As String

'Private mSig As PVersionInfo             '8
'Private mSignature As Long               '4
'Private mStrucVersionl As Integer        '2
'Private mStrucVersionh As Integer        '2
'Private mFileVersion As PVersionInfo     '8
'Private mProductVersion As PVersionInfo  '8
'Private mFileFlagsMask As Long           '4
'Private mFileFlags As Long               '4
'Private mFileOS As Long                  '4
'Private mFileType As Long                '4
'Private mFileSubtype As Long             '4
'Private mFileDateMS As Long              '4
'Private mFileDateLS As Long              '4
'                                    Sum 52

Private mComments As String
Private mCompanyName As String
Private mFileDescription As String

Private mFileMinor As Integer
Private mFileMajor As Integer
Private mFilePrivate As Integer
Private mFileBuild As Integer
Private mProductMinor As Integer
Private mProductMajor As Integer
Private mProductPrivate As Integer
Private mProductBuild As Integer
Private mFileFlags As Long

Private mFileName As String
Private mFileVersion As String
Private mInternalName As String
Private mLanguage As String
Private mLegalCopyright As String
Private mLegalTrademarks As String
Private mOriginalFilename As String
Private mPrivateBuild As String
Private mProductName As String
Private mProductVersion As String
Private mSpecialBuild As String

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  '
End Sub

'der Konstruktor nur Private, warum eigentlich? das weiß kein Mensch
Friend Sub NewC(StrFileName As String)
  'mStrPFN = StrFileName
  mFileName = StrFileName
End Sub
'
Friend Function New_FileVersionInfo(StrFileName As String) As FileVersionInfo
  Set New_FileVersionInfo = New FileVersionInfo
  Call New_FileVersionInfo.NewC(StrFileName)
End Function

'Private Shared Function ConvertTo8DigitHex(ByVal value As Integer) As String
Friend Function ConvertTo8DigitHex(ByVal value As Long) As String
  Dim StrVal As String: StrVal = Hex$(value)
  ConvertTo8DigitHex = String$(8 - Len(StrVal), "0") & StrVal
End Function

'Private Shared Function GetVarEntry(ByVal memPtr As IntPtr) As Integer
'Private Function GetVarEntry(ByVal memPtr As Long) As Long
Friend Function GetVarEntry(ByVal memPtr As IntPtr) As Long
  Dim num1 As Long
  'Dim ptr1 As Int32HiLo 'Long 'IntPtr = IntPtr.Zero
  Dim ptr1 As IntPtr: Set ptr1 = IntPtr.Zero
  If UnsafeNativeMethods.VerQueryValue(New_HandleRef(Nothing, memPtr), "\VarFileInfo\Translation", ptr1, num1) Then
    'Return ((Marshal.ReadInt16(ptr1) << 16) + Marshal.ReadInt16(DirectCast((CLng(ptr1) + 2), IntPtr)))
    GetVarEntry = Math.ShL(CLng(Marshal.ReadInt16(ptr1)), 16) + CLng(Marshal.ReadInt16(ptr1, 2))
    Exit Function
  End If
  'Return 67699940
  GetVarEntry = &H40904E4  '67699940
End Function

'Private Shared Function GetFileVersionLanguage(ByVal memPtr As IntPtr) As String
Friend Function GetFileVersionLanguage(ByVal memPtr As IntPtr) As String
  'Dim num1 As Integer = (FileVersionInfo.GetVarEntry(memPtr) >> 16)
  Dim num1 As Long: num1 = Math.ShR(FileVersionInfo.GetVarEntry(memPtr), 16)
  '16 bit nach rechts verschoben
  Dim builder1 As StringBuilder: Set builder1 = New_StringBuilder(, , , 256)
  'Dim buf As String * 256
  Call UnsafeNativeMethods.VerLanguageName(num1, builder1, builder1.Capacity)
  'Return builder1.ToString
  GetFileVersionLanguage = builder1.ToString 'Left$(buf, lstrlen(buf))
  'GetFileVersionLanguage = StrConv(GetFileVersionLanguage, vbUnicode)
End Function

'Private Shared Function GetFileVersionString(ByVal memPtr As IntPtr, ByVal name As String) As String
Friend Function GetFileVersionString(ByVal memPtr As IntPtr, ByVal name As String) As String
  Dim num1 As Long
  Dim text1 As String '= ""
  Dim ptr1 As IntPtr: Set ptr1 = IntPtr.Zero
  'Achtung VB-Bug:
  'If (UnsafeNativeMethods.VerQueryValue(New HandleRef(Nothing, memPtr), name, ptr1, num1) AndAlso (ptr1 <> IntPtr.Zero)) Then
  If (UnsafeNativeMethods.VerQueryValue(New_HandleRef(Nothing, memPtr), name, ptr1, num1)) Then
    If (ptr1 <> IntPtr.Zero) Then
      text1 = Marshal.PtrToStringAuto(ptr1)
    End If
  End If
  GetFileVersionString = text1
End Function

Friend Function GetVersionInfoForCodePage(ByVal memIntPtr As IntPtr, ByVal codepage As String) As Boolean
  Dim text1 As String: text1 = "\\StringFileInfo\\" '{0}\\{1}"
  mCompanyName = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "CompanyName")
  mFileDescription = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "FileDescription")
  mFileVersion = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "FileVersion")
  mInternalName = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "InternalName")
  mLegalCopyright = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "LegalCopyright")
  mOriginalFilename = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "OriginalFilename")
  mProductName = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "ProductName")
  mProductVersion = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "ProductVersion")
  mComments = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "Comments")
  mLegalTrademarks = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "LegalTrademarks")
  mPrivateBuild = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "PrivateBuild")
  mSpecialBuild = FileVersionInfo.GetFileVersionString(memIntPtr, text1 & codepage & "\\" & "SpecialBuild")
  mLanguage = FileVersionInfo.GetFileVersionLanguage(memIntPtr)
  Dim vs_fixedfileinfo1 As VS_FIXEDFILEINFO: Set vs_fixedfileinfo1 = FileVersionInfo.GetFixedFileInfo(memIntPtr)
  mFileMajor = FileVersionInfo.HIWORD(vs_fixedfileinfo1.dwFileVersionMS)
  mFileMinor = FileVersionInfo.LOWORD(vs_fixedfileinfo1.dwFileVersionMS)
  mFileBuild = FileVersionInfo.HIWORD(vs_fixedfileinfo1.dwFileVersionLS)
  mFilePrivate = FileVersionInfo.LOWORD(vs_fixedfileinfo1.dwFileVersionLS)
  mProductMajor = FileVersionInfo.HIWORD(vs_fixedfileinfo1.dwProductVersionMS)
  mProductMinor = FileVersionInfo.LOWORD(vs_fixedfileinfo1.dwProductVersionMS)
  mProductBuild = FileVersionInfo.HIWORD(vs_fixedfileinfo1.dwProductVersionLS)
  mProductPrivate = FileVersionInfo.LOWORD(vs_fixedfileinfo1.dwProductVersionLS)
  mFileFlags = vs_fixedfileinfo1.dwFileFlags
  'Return (Not Me.fileVersion Is String.Empty)
  GetVersionInfoForCodePage = Not (Len(mFileVersion) = 0)
End Function

'Private Shared Function GetFixedFileInfo(ByVal memPtr As IntPtr) As VS_FIXEDFILEINFO
Friend Function GetFixedFileInfo(ByVal memPtr As IntPtr) As VS_FIXEDFILEINFO
  Dim num1 As Long
  Dim ptr1 As IntPtr: Set ptr1 = IntPtr.Zero
  If UnsafeNativeMethods.VerQueryValue(New_HandleRef(Nothing, memPtr), "\", ptr1, num1) Then
    Dim vs_fixedfileinfo1 As New VS_FIXEDFILEINFO
    Call Marshal.PtrToStructure(ptr1, vs_fixedfileinfo1)
    'Return vs_fixedfileinfo1
    Set GetFixedFileInfo = vs_fixedfileinfo1: Exit Function
  End If
  'Return New VS_FIXEDFILEINFO
  Set GetFixedFileInfo = New VS_FIXEDFILEINFO
End Function

'Private Shared Function HIWORD(ByVal dword As Integer) As Integer
Friend Function HIWORD(ByVal n As Long) As Long
  'Return Util.HIWORD(dword)
  HIWORD = (Math.ShR(n, 16) And 65535)
  'oder auch:
  'HIWORD = (n \ 65536) And &HFFFF&
End Function

'Private Shared Function LOWORD(ByVal dword As Integer) As Integer
Friend Function LOWORD(ByVal n As Long) As Long
  'Return Util.LOWORD(dword)
  LOWORD = (n And 65535)
  'oder auch:
  'LOWORD = n And &HFFFF&
End Function



'Public Shared Function GetVersionInfo(ByVal fileName As String) As System.Diagnostics.FileVersionInfo
'Shared also entweder in ein Modul, oder
'ein globales Objekt der Klasse halten
'Also am besten in ein Modul
Public Function GetVersionInfo(ByVal FileName As String) As FileVersionInfo
  Dim num1 As Long
  If Not (LenB(Dir$(FileName)) <> 0) Then
    'Dim text1 As String = FileVersionInfo.GetFullPathWithAssert(fileName)
    'New FileIOPermission(FileIOPermissionAccess.Read, text1).Demand
    'Throw New FileNotFoundException(fileName)
    MsgBox "FileNotFoundException: " & FileName
    Exit Function
  End If
  Dim num2 As Long: num2 = UnsafeNativeMethods.GetFileVersionInfoSize(FileName, num1)
  Dim info1 As FileVersionInfo: Set info1 = New_FileVersionInfo(FileName)
  If (num2 = 0) Then
    'Return info1
    Set GetVersionInfo = info1: Exit Function
  End If
  Dim buffer1() As Byte: ReDim buffer1(0 To num2 - 1)
  'Dim numRef1 As Byte*
  'Static numRef1 As Long ': numRef1 = VarPtr(buffer1(0))
  Dim numRef1 As Long
  'Fixed numRef1 = buffer1
  'Call Fixed(VarPtr(buffer1(0))): numRef1 = VarPtr(buffer1(0))
  'Call Fixed(buffer1)
    numRef1 = VarPtr(buffer1(0))
    
    Dim ptr1 As IntPtr: Set ptr1 = New_IntPtr(numRef1)
    If UnsafeNativeMethods.GetFileVersionInfo(FileName, 0, num2, New_HandleRef(Nothing, ptr1)) Then
      Debug.Print buffer1
      Dim num3 As Long: num3 = FileVersionInfo.GetVarEntry(ptr1)
      If Not info1.GetVersionInfoForCodePage(ptr1, FileVersionInfo.ConvertTo8DigitHex(num3)) Then
        Dim numArray1 As Variant: numArray1 = Array(67699888, 67699940, 67698688)
        '67699888 = &H040904B0
        '67699940 = &H040904E4
        '67698688 = &H04090000
        Dim num4 'As Integer
        For Each num4 In numArray1
          If ((num4 <> num3) And info1.GetVersionInfoForCodePage(ptr1, FileVersionInfo.ConvertTo8DigitHex(num4))) Then
            GoTo Label_00DB
          End If
        Next
      End If
    End If
  'End Fixed
  'Call Unfixed(buffer1)
  'Call Unfixed(ArrPtr(buffer1))
Label_00DB:
  'Return info1
  Set GetVersionInfo = info1
End Function



'Public Function GetVersionInfo(Optional StrFileName As String) As FileVersionInfo
'Dim SVer As String, FVer As String, PVer As String
'Dim F As String
'Dim Typ As String, STyp As String, OS As String
'Dim Buff() As Byte
'Dim L As Long, BuffL As Long, Pointer As Long
''Dim VersionInfo As VS_FIXEDFILEINFO
'  If Len(StrFileName) > 0 Then mStrPFN = StrFileName
'  If Len(mStrPFN) > 0 Then
'
'    L = GetFileVersionInfoSizeA(mStrPFN, 0&)
'    If L < 1 Then
'      'MsgBox "Keine Versions-Info vorhanden!"
'    Else
'      ReDim Buff(0 To L)
'      Call GetFileVersionInfoA(mStrPFN, 0&, L, Buff(0))
'      Call VerQueryValueA(Buff(0), "\", Pointer, BuffL)
'      Call RtlMoveMemory(ByVal VarPtr(mSignature), ByVal (Pointer), 52) 'LenB(Version))
'
''      mStrFileVersion11 = Format$(Version.dwFileVersionMSh)
''      mStrFileVersion12 = Format$(Version.dwFileVersionMSl)
''      mStrFileVersion21 = Format$(Version.dwFileVersionLSh)
''      mStrFileVersion22 = Format$(Version.dwFileVersionLSl)
''
''      mStrProdVersion11 = Format$(Version.dwProductVersionMSh)
''      mStrProdVersion12 = Format$(Version.dwProductVersionMSl)
''      mStrProdVersion21 = Format$(Version.dwProductVersionLSh)
''      mStrProdVersion22 = Format$(Version.dwProductVersionLSl)
''
''      mStrStrucVersion1 = Format$(Version.dwStrucVersionh)
''      mStrStrucVersion2 = Format$(Version.dwStrucVersionl)
'
''      With Version
''        SVer = Format$(.dwStrucVersionh) & "." & _
''               Format$(.dwStrucVersionl)
''
''        FVer = Format$(.dwFileVersionMSh) & "." & _
''               Format$(.dwFileVersionMSl) & "." & _
''               Format$(.dwFileVersionLSh) & "." & _
''               Format$(.dwFileVersionLSl)
''
''        PVer = Format$(.dwProductVersionMSh) & "." & _
''               Format$(.dwProductVersionMSl) & "." & _
''               Format$(.dwProductVersionLSh) & "." & _
''               Format$(.dwProductVersionLSl)
'
'        If mFileFlags And VS_FF_DEBUG Then F = "Debug "
'        If mFileFlags And VS_FF_PRERELEASE Then F = F & "PreRel "
'        If mFileFlags And VS_FF_PATCHED Then F = F & "Patched "
'        If mFileFlags And VS_FF_PRIVATEBUILD Then F = F & "Private "
'        If mFileFlags And VS_FF_INFOINFERRED Then F = F & "Info "
'        If mFileFlags And VS_FF_SPECIALBUILD Then F = F & "Special "
'        If mFileFlags And VFT2_UNKNOWN Then F = F & "Unknown "
'      'End With
'
'      Select Case mFileOS
'        Case VOS_DOS_WINDOWS16: OS = "DOS-Win16"
'        Case VOS_DOS_WINDOWS32: OS = "DOS-Win32"
'        Case VOS_OS216_PM16:    OS = "OS/2-16 PM-16"
'        Case VOS_OS232_PM32:    OS = "OS/2-16 PM-32"
'        Case VOS_NT_WINDOWS32:  OS = "NT-Win32"
'        Case Else:              OS = "Unbekannt"
'      End Select
'
'      Select Case mFileType
'        Case VFT_APP:                Typ = "App"
'        Case VFT_DLL:                Typ = "DLL"
'        Case VFT_DRV:                Typ = "Driver"
'          Select Case mFileSubtype
'            Case VFT2_DRV_PRINTER:     STyp = "Printer drv"
'            Case VFT2_DRV_KEYBOARD:    STyp = "Keyboard drv"
'            Case VFT2_DRV_LANGUAGE:    STyp = "Language drv"
'            Case VFT2_DRV_DISPLAY:     STyp = "Display drv"
'            Case VFT2_DRV_MOUSE:       STyp = "Mouse drv"
'            Case VFT2_DRV_NETWORK:     STyp = "Network drv"
'            Case VFT2_DRV_SYSTEM:      STyp = "System drv"
'            Case VFT2_DRV_INSTALLABLE: STyp = "Installable"
'            Case VFT2_DRV_SOUND:       STyp = "Sound drv"
'            Case VFT2_DRV_COMM:        STyp = "Comm drv"
'            Case VFT2_UNKNOWN:         STyp = "Unknown"
'          End Select
'        Case VFT_FONT:               Typ = "Font"
'          Select Case mFileSubtype
'            Case VFT2_FONT_RASTER:     STyp = "Raster Font"
'            Case VFT2_FONT_VECTOR:     STyp = "Vector Font"
'            Case VFT2_FONT_TRUETYPE:   STyp = "TrueType Font"
'          End Select
'        Case VFT_VXD:                Typ = "VxD"
'        Case VFT_STATIC_LIB:         Typ = "Lib"
'        Case Else:                   Typ = "Unbekannt"
'      End Select
'      'select case version.dwFileDateLS
'    End If
'  End If
'
'  Set GetVersionInfo = Me
'End Function

'Alle Properties sind ReadOnly
'Public ReadOnly Property FileName() As String
'Public Property Get fileName() As String: fileName = mStrPFN: End Property
Public Property Get FileName() As String: FileName = mFileName: End Property

'Public ReadOnly Property InternalName() As String
Public Property Get InternalName() As String
  InternalName = mInternalName
End Property

'Public ReadOnly Property OriginalFilename() As String
Public Property Get OriginalFilename() As String
  OriginalFilename = mOriginalFilename
End Property

'Public ReadOnly Property Comments() As String
Public Property Get Comments() As String
  Comments = mComments
End Property

'Public ReadOnly Property CompanyName() As String
Public Property Get CompanyName() As String
  CompanyName = mCompanyName
End Property

'Public ReadOnly Property Language() As String
Public Property Get Language() As String
  Language = mLanguage
End Property

'Public ReadOnly Property LegalCopyright() As String
Public Property Get LegalCopyright() As String
  LegalCopyright = mLegalCopyright
End Property
'Public ReadOnly Property LegalTrademarks() As String
Public Property Get LegalTrademarks() As String
  LegalTrademarks = mLegalTrademarks
End Property

'Public ReadOnly Property PrivateBuild() As String
Public Property Get PrivateBuild() As String
  PrivateBuild = mPrivateBuild
End Property

'Public ReadOnly Property SpecialBuild() As String
Public Property Get SpecialBuild() As String
  SpecialBuild = mSpecialBuild
End Property


'##############################'    Fileversion    '##############################'
'Public ReadOnly Property FileVersion() As String
Public Property Get FileVersion() As String
'  With mFileVersion
'    FileVersion = CStr(.mMajor) & "." & CStr(.mMinor) & "." & CStr(.mBuild) & "." & CStr(.mPrivate)
'  End With
  FileVersion = mFileVersion
End Property
''Public ReadOnly Property FileMajorPart() As Integer
'Public Property Get FileMajorPart() As Long: FileMajorPart = CLng(mFileVersion.mMajor): End Property
''Public ReadOnly Property FileMinorPart() As Integer
'Public Property Get FileMinorPart() As Long: FileMinorPart = CLng(mFileVersion.mMinor): End Property
''Public ReadOnly Property FileBuildPart() As Integer
'Public Property Get FileBuildPart() As Long: FileBuildPart = CLng(mFileVersion.mBuild): End Property
''Public ReadOnly Property FilePrivatePart() As Integer
'Public Property Get FilePrivatePart() As Long: FilePrivatePart = CLng(mFileVersion.mPrivate): End Property

'Public ReadOnly Property FileMajorPart() As Integer
Public Property Get FileMajorPart() As Long: FileMajorPart = CLng(mFileMajor): End Property
'Public ReadOnly Property FileMinorPart() As Integer
Public Property Get FileMinorPart() As Long: FileMinorPart = CLng(mFileMinor): End Property
'Public ReadOnly Property FileBuildPart() As Integer
Public Property Get FileBuildPart() As Long: FileBuildPart = CLng(mFileBuild): End Property
'Public ReadOnly Property FilePrivatePart() As Integer
Public Property Get FilePrivatePart() As Long: FilePrivatePart = CLng(mFilePrivate): End Property



'------------------------------------------------------------------
'Public ReadOnly Property FileDescription() As String
Public Property Get FileDescription() As String
  FileDescription = mFileDescription
End Property
'Public ReadOnly Property ProductName() As String
Public Property Get ProductName() As String
  ProductName = mProductName
End Property
'------------------------------------------------------------------

'##############################'    Productversion    '##############################'
'Public ReadOnly Property ProductVersion() As String
Public Property Get ProductVersion() As String
'  With mProductVersion
'    productVersion = CStr(.mMajor) & "." & CStr(.mMinor) & "." & CStr(.mBuild) & "." & CStr(.mPrivate)
'  End With
  ProductVersion = mProductVersion
End Property
''Public ReadOnly Property ProductMajorPart() As Integer
'Public Property Get ProductMajorPart() As Long: ProductMajorPart = CLng(mProductVersion.mMajor): End Property
''Public ReadOnly Property ProductMinorPart() As Integer
'Public Property Get ProductMinorPart() As Long: ProductMinorPart = CLng(mProductVersion.mMajor): End Property
''Public ReadOnly Property ProductBuildPart() As Integer
'Public Property Get ProductBuildPart() As Long: ProductBuildPart = CLng(mProductVersion.mMajor): End Property
''Public ReadOnly Property ProductPrivatePart() As Integer
'Public Property Get ProductPrivatePart() As Long: ProductPrivatePart = CLng(mProductVersion.mMajor): End Property

'Public ReadOnly Property ProductMajorPart() As Integer
Public Property Get ProductMajorPart() As Long: ProductMajorPart = CLng(mProductMajor): End Property
'Public ReadOnly Property ProductMinorPart() As Integer
Public Property Get ProductMinorPart() As Long: ProductMinorPart = CLng(mProductMinor): End Property
'Public ReadOnly Property ProductBuildPart() As Integer
Public Property Get ProductBuildPart() As Long: ProductBuildPart = CLng(mProductBuild): End Property
'Public ReadOnly Property ProductPrivatePart() As Integer
Public Property Get ProductPrivatePart() As Long: ProductPrivatePart = CLng(mProductPrivate): End Property





'##############################'    Booleans     '##############################'
'Public ReadOnly Property IsDebug() As Boolean
Public Property Get IsDebug() As Boolean
  IsDebug = ((mFileFlags And 1) <> 0)
End Property
'Public ReadOnly Property IsPatched() As Boolean
Public Property Get IsPatched() As Boolean
  IsPatched = ((mFileFlags And 4) <> 0)
End Property
'Public ReadOnly Property IsPreRelease() As Boolean
Public Property Get IsPreRelease() As Boolean
  IsPreRelease = ((mFileFlags And 2) <> 0)
End Property
'Public ReadOnly Property IsPrivateBuild() As Boolean
Public Property Get IsPrivateBuild() As Boolean
  IsPrivateBuild = ((mFileFlags And 8) <> 0)
End Property
'Public ReadOnly Property IsSpecialBuild() As Boolean
Public Property Get IsSpecialBuild() As Boolean
  IsSpecialBuild = ((mFileFlags And 32) <> 0)
End Property


'Public Overrides Function ToString() As String
Public Function ToString() As String
'Zusammenfassung:
' Gibt eine unvollständige Liste der Eigenschaften in
' System.Diagnostics.FileVersionInfo und deren Werte zurück.
'Rückgabewerte:
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
'Wenn keine Datei mit dem angegebenen Namen gefunden wurde, enthält diese
'Liste nur den Namen der angeforderten Datei. Boolesche Werte sind auf false
'und alle anderen Einträge auf null festgelegt.
  Dim SB As StringBuilder: Set SB = New_StringBuilder(, , , 128)
  Call SB.Append("File:             ").Append(mFileName).Append(vbCrLf)
  Call SB.Append("InternalName:     ").Append(mInternalName).Append(vbCrLf)
  Call SB.Append("OriginalFilename: ").Append(mOriginalFilename).Append(vbCrLf)
  Call SB.Append("FileVersion:      ").Append(mFileVersion).Append(vbCrLf)
  Call SB.Append("FileDescription:  ").Append(mFileDescription).Append(vbCrLf)
  Call SB.Append("Product:          ").Append(mProductName).Append(vbCrLf)
  Call SB.Append("ProductVersion:   ").Append(mProductVersion).Append(vbCrLf)
  Call SB.Append("Debug:            ").Append(CStr(Me.IsDebug)).Append(vbCrLf)
  Call SB.Append("Patched:          ").Append(CStr(Me.IsPatched)).Append(vbCrLf)
  Call SB.Append("PreRelease:       ").Append(CStr(Me.IsPreRelease)).Append(vbCrLf)
  Call SB.Append("PrivateBuild:     ").Append(CStr(Me.IsPrivateBuild)).Append(vbCrLf)
  Call SB.Append("SpecialBuild:     ").Append(CStr(Me.IsSpecialBuild)).Append(vbCrLf)
  Call SB.Append("Language:         ").Append(Me.Language).Append(vbCrLf)
  ToString = SB.ToString
'
'if you have StringBuilder anyway, and need for speed, ok use the upper part.
'well, but we can make it also a bit more independent from other stuff, so use
'the part below here:
'  ToString = ToString & "File:             " & Me.FileName & vbCrLf
'  ToString = ToString & "InternalName:     " & Me.InternalName & vbCrLf
'  ToString = ToString & "OriginalFilename: " & Me.OriginalFilename & vbCrLf
'  ToString = ToString & "FileVersion:      " & Me.FileVersion & vbCrLf
'  ToString = ToString & "FileDescription:  " & Me.FileDescription & vbCrLf
'  ToString = ToString & "Product:          " & Me.ProductName & vbCrLf
'  ToString = ToString & "ProductVersion:   " & Me.ProductVersion & vbCrLf
'  ToString = ToString & "Debug:            " & CStr(Me.IsDebug) & vbCrLf
'  ToString = ToString & "Patched:          " & CStr(Me.IsPatched) & vbCrLf
'  ToString = ToString & "PreRelease:       " & CStr(Me.IsPreRelease) & vbCrLf
'  ToString = ToString & "PrivateBuild:     " & CStr(Me.IsPrivateBuild) & vbCrLf
'  ToString = ToString & "SpecialBuild:     " & CStr(Me.IsSpecialBuild) & vbCrLf
'  ToString = ToString & "Language:         " & Me.Language & vbCrLf

End Function

'Public Function GetStrFileVersion(Optional StrPath As String) As String
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetStrFileVersion = mStrFileVersion11 & "." & mStrFileVersion12 & "." & mStrFileVersion21 & "." & mStrFileVersion22
'End Function
'
'Public Function GetDblFileVersion(Optional StrPath As String) As Double
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetDblFileVersion = Val(mStrFileVersion11 & "." & mStrFileVersion12)
'End Function
'
'Public Function GetStrProdVersion(Optional StrPath As String) As String
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetStrProdVersion = mStrProdVersion11 & "." & mStrProdVersion12 & "." & mStrProdVersion21 & "." & mStrProdVersion22
'End Function
'
'Public Function GetDblProdVersion(Optional StrPath As String) As Double
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetDblProdVersion = Val(mStrProdVersion11 & "." & mStrProdVersion12)
'End Function
'
'Public Function GetStrStrucVersion(Optional StrPath As String) As String
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetStrStrucVersion = mStrStrucVersion1 & "." & mStrStrucVersion2
'End Function
'
'Public Function GetDblStrucVersion(Optional StrPath As String) As Double
'  If StrPath <> "" Then mStrPathFileName = StrPath
'  Call GetVersion(mStrPathFileName)
'  GetDblStrucVersion = Val(mStrStrucVersion1 & "." & mStrStrucVersion2)
'End Function




