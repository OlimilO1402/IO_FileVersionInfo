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
#If VBA7 Then
    'we assume in VBA7 ansi is obsolete
    Private Declare PtrSafe Function GetFileVersionInfoSize Lib "version" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As LongPtr, ByRef lpdwHandle As LongPtr) As Long
    Private Declare PtrSafe Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As LongPtr, ByVal dwHandle As LongPtr, ByVal dwLen As Long, lpData As Any) As Long
    Private Declare PtrSafe Function VerQueryValue Lib "version" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As LongPtr, lplpBuffer As Any, puLen As Long) As Long
    Private Declare PtrSafe Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameW" (ByVal wLang As Long, ByVal szLang As LongPtr, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)
#Else
Private Enum LongPtr
    [_]
End Enum
#Const defUnicode = 1
#If defUnicode Then
    Private Declare Function GetFileVersionInfoSize Lib "version" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As LongPtr, ByRef lpdwHandle As LongPtr) As Long
    Private Declare Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As LongPtr, ByVal dwHandle As LongPtr, ByVal dwLen As Long, lpData As Any) As Long
    Private Declare Function VerQueryValue Lib "version" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As LongPtr, lplpBuffer As Any, puLen As Long) As Long
    Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameW" (ByVal wLang As Long, ByVal szLang As LongPtr, ByVal nSize As Long) As Long
    Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
    Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal dst As LongPtr, ByVal src As LongPtr) As Long
#Else
    Private Declare Function GetFileVersionInfoSize Lib "version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As LongPtr, ByRef lpdwHandle As LongPtr) As Long
    Private Declare Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As LongPtr, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
    Private Declare Function VerQueryValue Lib "version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As LongPtr, lplpBuffer As Any, puLen As Long) As Long
    Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As LongPtr, ByVal nSize As Long) As Long
    Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As LongPtr) As Long
    Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal dst As LongPtr, ByVal src As LongPtr) As Long
#End If
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLen As Long)

#End If

Private Const VS_FFI_FILEFLAGSMASK As Long = &H3F&
Private Const VS_FFI_STRUCVERSION  As Long = &H10000
Private Const VS_FFI_SIGNATURE     As Long = &HFEEF04BD

Private Const VS_FF_DEBUG          As Long = &H1&
Private Const VS_FF_PRERELEASE     As Long = &H2&
Private Const VS_FF_PATCHED        As Long = &H4&
Private Const VS_FF_PRIVATEBUILD   As Long = &H8&
Private Const VS_FF_INFOINFERRED   As Long = &H10&
Private Const VS_FF_SPECIALBUILD   As Long = &H20&

Private Const VOS_UNKNOWN          As Long = &H0&
Private Const VOS_DOS              As Long = &H10000
Private Const VOS_OS216            As Long = &H20000
Private Const VOS_OS232            As Long = &H30000
Private Const VOS_NT               As Long = &H40000
Private Const VOS__BASE            As Long = &H0&
Private Const VOS__WINDOWS16       As Long = &H1&
Private Const VOS__PM16            As Long = &H2&
Private Const VOS__PM32            As Long = &H3&
Private Const VOS__WINDOWS32       As Long = &H4&

Private Const VOS_DOS_WINDOWS16    As Long = &H10001
Private Const VOS_DOS_WINDOWS32    As Long = &H10004
Private Const VOS_OS216_PM16       As Long = &H20002
Private Const VOS_OS232_PM32       As Long = &H30003
Private Const VOS_NT_WINDOWS32     As Long = &H40004

Private Const VFT_UNKNOWN          As Long = &H0&
Private Const VFT_APP              As Long = &H1&
Private Const VFT_DLL              As Long = &H2&
Private Const VFT_DRV              As Long = &H3&
Private Const VFT_FONT             As Long = &H4&
Private Const VFT_VXD              As Long = &H5&
Private Const VFT_STATIC_LIB       As Long = &H7&

Private Const VFT2_UNKNOWN         As Long = &H0&
Private Const VFT2_DRV_PRINTER     As Long = &H1&
Private Const VFT2_DRV_KEYBOARD    As Long = &H2&
Private Const VFT2_DRV_LANGUAGE    As Long = &H3&
Private Const VFT2_DRV_DISPLAY     As Long = &H4&
Private Const VFT2_DRV_MOUSE       As Long = &H5&
Private Const VFT2_DRV_NETWORK     As Long = &H6&
Private Const VFT2_DRV_SYSTEM      As Long = &H7&
Private Const VFT2_DRV_INSTALLABLE As Long = &H8&
Private Const VFT2_DRV_SOUND       As Long = &H9&
Private Const VFT2_DRV_COMM        As Long = &HA&

Private Const VFT2_FONT_RASTER     As Long = &H1&
Private Const VFT2_FONT_VECTOR     As Long = &H2&
Private Const VFT2_FONT_TRUETYPE   As Long = &H3&

Private Const MAX_PATH             As Long = 260

'dw = doubleword = int32 = Long
'the versions here are only words so drop the "d"
Private Type VS_FIXEDFILEINFO
    dwSignature        As Long
    wStrucVersionl     As Integer ' minor      'unsigned!
    wStrucVersionh     As Integer 'major       'unsigned!
    wFileVersionMSl    As Integer ' minor      'unsigned!
    wFileVersionMSh    As Integer 'major       'unsigned!
    wFileVersionLSl    As Integer '   private  'unsigned!
    wFileVersionLSh    As Integer '  build     'unsigned!
    wProductVersionMSl As Integer ' minor      'unsigned!
    wProductVersionMSh As Integer 'major       'unsigned!
    wProductVersionLSl As Integer '   private  'unsigned!
    wProductVersionLSh As Integer '  build     'unsigned!
    dwFileFlagsMask    As Long
    dwFileFlags        As Long
    dwFileOS           As Long
    dwFileType         As Long
    dwFileSubtype      As Long
    dwFileDateMS       As Long
    dwFileDateLS       As Long
End Type
Private mFileName         As String
Private mVSFileInfo       As VS_FIXEDFILEINFO
Private mComments         As String
Private mCompanyName      As String
Private mFileDescription  As String
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

Friend Sub New_(aPathFileName As String)
Try: On Error GoTo Catch
    mFileName = aPathFileName
    If (LenB(Dir$(mFileName)) = 0) Then
        MsgBox "FileNotFoundException: " & mFileName
        Exit Sub
    End If
    Dim siz As Long: siz = GetFileVersionInfoSize(StrPtr(mFileName), 0)
    If (siz = 0) Then
        'Set GetVersionInfo = info1:
    End If
    ReDim Buffer(0 To siz - 1) As Byte
    Dim pBuffer  As LongPtr: pBuffer = VarPtr(Buffer(0))
    If CBool(GetFileVersionInfo(StrPtr(mFileName), 0, siz, ByVal pBuffer)) Then
        'Debug.Print buffer1
        Dim ve As Long: ve = GetVarEntry(pBuffer) 'num3
        'If Not info1.GetVersionInfoForCodePage(ptr1, ConvertTo8DigitHex(num3)) Then
        If Not GetVersionInfoForCodePage(pBuffer, ConvertTo8DigitHex(ve)) Then
            'Wenn die Sprache nicht geklappt hat, dann noch mit den drei anderen probieren
            Dim v, vArr: vArr = Array(&H40904B0, &H40904E4, &H4090000) 'num4
            For Each v In vArr
                If (v <> ve) Then
                    'If info1.GetVersionInfoForCodePage(ptr1, ConvertTo8DigitHex(num4)) Then
                    If GetVersionInfoForCodePage(pBuffer, ConvertTo8DigitHex(v)) Then
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    'Set GetVersionInfo = info1
    Exit Sub
Catch:
    MsgBox "Error in creating FileVersionInfo"
End Sub

'in ein Modul MNew kopieren
Function FileVersionInfo(aPathFileName As String) As FileVersionInfo
    Set FileVersionInfo = New FileVersionInfo: FileVersionInfo.New_ aPathFileName
End Function

'Private Shared Function ConvertTo8DigitHex(ByVal value As Integer) As String
Friend Function ConvertTo8DigitHex(ByVal value As Long) As String
    Dim s As String: s = Hex$(value)
    ConvertTo8DigitHex = String$(8 - Len(s), "0") & s
End Function

'Private Shared Function GetVarEntry(ByVal memPtr As IntPtr) As Integer
Friend Function GetVarEntry(ByVal memPtr As LongPtr) As Long
    Dim num1 As Long, ptr1 As LongPtr
    If VerQueryValue(ByVal memPtr, StrPtr("\VarFileInfo\Translation"), ByVal VarPtr(ptr1), num1) Then
        GetVarEntry = ShL(ReadInt16(ptr1), 16) + CLng(ReadInt16(ptr1, 2))
        Exit Function
    End If
    GetVarEntry = &H40904E4  '67699940
End Function

'Private Shared Function GetFileVersionLanguage(ByVal memPtr As IntPtr) As String
Friend Function GetFileVersionLanguage(ByVal memPtr As LongPtr) As String
    Dim num1 As Long: num1 = ShR(GetVarEntry(memPtr), 16)
    'Dim buffer As String * 256
    ReDim Buffer(0 To 511) As Byte
    Dim l As Long: l = VerLanguageName(num1, VarPtr(Buffer(0)), 256)  'hmm hier * 2???????????????????????
    GetFileVersionLanguage = Left$(Buffer, l)
End Function

'Private Shared Function GetFileVersionString(ByVal memPtr As IntPtr, ByVal name As String) As String
Friend Function GetFileVersionString(ByVal memPtr As LongPtr, ByVal name As String) As String
    Dim num1 As Long, ptr1 As LongPtr
    If VerQueryValue(ByVal memPtr, StrPtr(name), ptr1, num1) Then
        If (ptr1 <> 0) Then
            GetFileVersionString = PtrToString(ptr1)
        End If
    End If
End Function

Friend Function GetVersionInfoForCodePage(ByVal memIntPtr As LongPtr, ByVal codepage As String) As Boolean
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
Friend Function GetFixedFileInfo(ByVal memPtr As LongPtr) As VS_FIXEDFILEINFO
    Dim num1 As Long, ptr As LongPtr
    VerQueryValue ByVal memPtr, StrPtr("\"), ptr, num1
    PtrToStructure ptr, VarPtr(GetFixedFileInfo), LenB(GetFixedFileInfo)
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
Private Function HIWORD(ByVal n As Long) As Long
    'HIWORD = (Math.ShR(n, 16) And 65535)
    'oder auch:
    HIWORD = (n \ 65536) And &HFFFF&
End Function
'Private Shared Function LOWORD(ByVal dword As Integer) As Integer
Private Function LOWORD(ByVal n As Long) As Long
    'LOWORD = (n And 65535)
    'oder auch:
    LOWORD = n And &HFFFF&
End Function

'##############################'  Marshal   '##############################'
Private Function ReadInt16(ByVal ptr As LongPtr, Optional ByVal ofs As Long) As Integer
    '                   pDst             , pSrc
    RtlMoveMemory ByVal VarPtr(ReadInt16), ByVal (ptr + ofs), 2
End Function
Private Function PtrToString(ByVal pStr As LongPtr, Optional ByVal sLen As Long) As String
    If (pStr = 0) Then
        MsgBox "Marshal.PtrToString: pStr=Nullpointer"
        Exit Function
    End If
    Dim l As Long: l = lstrlen(pStr)
    PtrToString = Space$(l)
    lstrcpy StrPtr(PtrToString), pStr
#If defUnicode Then
    'ist es dann schon der richtige String?
    'MsgBox PtrToString
#Else
    PtrToString = Left$(StrConv(PtrToString, vbUnicode), num1)
#End If
End Function

'Public Shared Function PtrToStructure(ByVal ptr As System.IntPtr, ByVal structureType As System.Type) As Object
Private Sub PtrToStructure(ByVal ptr As LongPtr, ByVal pStruct As LongPtr, ByVal LenBStruct As Long)
    'hier wird nicht der Pointer übertragen, sondern vielmehr der gesamte Speicherbereich
    'in das Objekt structure hineinkopiert
    RtlMoveMemory ByVal pStruct, ByVal ptr, LenBStruct
End Sub

'##############################'   My Properties   '##############################'
'All properties ReadOnly
Public Property Get FileName() As String:                 FileName = mFileName:         End Property
Public Property Get InternalName() As String:         InternalName = mInternalName:     End Property
Public Property Get OriginalFilename() As String: OriginalFilename = mOriginalFilename: End Property
Public Property Get Comments() As String:                 Comments = mComments:         End Property
Public Property Get CompanyName() As String:           CompanyName = mCompanyName:      End Property
Public Property Get Language() As String:                 Language = mLanguage:         End Property
Public Property Get LegalCopyright() As String:     LegalCopyright = mLegalCopyright:   End Property
Public Property Get LegalTrademarks() As String:   LegalTrademarks = mLegalTrademarks:  End Property
Public Property Get PrivateBuild() As String:         PrivateBuild = mPrivateBuild:     End Property
Public Property Get SpecialBuild() As String:         SpecialBuild = mSpecialBuild:     End Property

'##############################'    Fileversion    '##############################'
Public Property Get FileVersion() As String:       FileVersion = mFileVersion:                      End Property
Public Property Get FileMajorPart() As Long:     FileMajorPart = UInt16_ToInt32(mVSFileInfo.wFileVersionMSh): End Property
Public Property Get FileMinorPart() As Long:     FileMinorPart = UInt16_ToInt32(mVSFileInfo.wFileVersionMSl): End Property
Public Property Get FileBuildPart() As Long:     FileBuildPart = UInt16_ToInt32(mVSFileInfo.wFileVersionLSh): End Property
Public Property Get FilePrivatePart() As Long: FilePrivatePart = UInt16_ToInt32(mVSFileInfo.wFileVersionLSl): End Property

'------------------------------------------------------------------
Public Property Get FileDescription() As String:  FileDescription = mFileDescription: End Property
Public Property Get ProductName() As String:  ProductName = mProductName: End Property
'------------------------------------------------------------------

'##############################'  Productversion   '##############################'
Public Property Get ProductVersion() As String:       ProductVersion = mProductVersion:                      End Property
Public Property Get ProductMajorPart() As Long:     ProductMajorPart = UInt16_ToInt32(mVSFileInfo.wProductVersionMSh): End Property
Public Property Get ProductMinorPart() As Long:     ProductMinorPart = UInt16_ToInt32(mVSFileInfo.wProductVersionMSl): End Property
Public Property Get ProductBuildPart() As Long:     ProductBuildPart = UInt16_ToInt32(mVSFileInfo.wProductVersionLSh): End Property
Public Property Get ProductPrivatePart() As Long: ProductPrivatePart = UInt16_ToInt32(mVSFileInfo.wProductVersionLSl): End Property

Private Function UInt16_ToInt32(ByVal uint16 As Integer) As Long
    If uint16 < 0 Then UInt16_ToInt32 = 65536 + uint16 Else UInt16_ToInt32 = uint16
End Function

'##############################'    Bool Props     '##############################'
'All properties ReadOnly
Public Property Get IsDebug() As Boolean:                IsDebug = ((mVSFileInfo.dwFileFlags And VS_FF_DEBUG) <> 0):        End Property
Public Property Get IsPreRelease() As Boolean:      IsPreRelease = ((mVSFileInfo.dwFileFlags And VS_FF_PRERELEASE) <> 0):   End Property
Public Property Get IsPatched() As Boolean:            IsPatched = ((mVSFileInfo.dwFileFlags And VS_FF_PATCHED) <> 0):      End Property
Public Property Get IsPrivateBuild() As Boolean:  IsPrivateBuild = ((mVSFileInfo.dwFileFlags And VS_FF_PRIVATEBUILD) <> 0): End Property
Public Property Get IsInfoInferred() As Boolean:  IsInfoInferred = ((mVSFileInfo.dwFileFlags And VS_FF_INFOINFERRED) <> 0): End Property
Public Property Get IsSpecialBuild() As Boolean:  IsSpecialBuild = ((mVSFileInfo.dwFileFlags And VS_FF_SPECIALBUILD) <> 0): End Property

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
Public Function ToStr() As String
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
    Dim s As String
    s = s & "File:             " & mFileName & vbCrLf
    s = s & "InternalName:     " & mInternalName & vbCrLf
    s = s & "OriginalFilename: " & mOriginalFilename & vbCrLf
    s = s & "FileVersion:      " & mFileVersion & vbCrLf
    s = s & "FileDescription:  " & mFileDescription & vbCrLf
    s = s & "Product:          " & mProductName & vbCrLf
    s = s & "ProductVersion:   " & mProductVersion & vbCrLf
    s = s & "Debug:            " & CStr(Me.IsDebug) & vbCrLf
    s = s & "Patched:          " & CStr(Me.IsPatched) & vbCrLf
    s = s & "PreRelease:       " & CStr(Me.IsPreRelease) & vbCrLf
    s = s & "PrivateBuild:     " & CStr(Me.IsPrivateBuild) & vbCrLf
    s = s & "SpecialBuild:     " & CStr(Me.IsSpecialBuild) & vbCrLf
    s = s & "Language:         " & Me.Language & vbCrLf
    s = s & "InfoInferred: " & CStr(Me.IsInfoInferred) & vbCrLf
    s = s & "FileVInt:     " & Me.FileMajorPart & "." & Me.FileMinorPart & "." & Me.FileBuildPart & "." & Me.FilePrivatePart & vbCrLf
    s = s & "ProdVInt:     " & Me.ProductMajorPart & "." & Me.ProductMinorPart & "." & Me.ProductBuildPart & "." & Me.ProductPrivatePart & vbCrLf
    s = s & "FileType:     " & Me.FileType & vbCrLf
    s = s & "FileOS:       " & Me.FileOS
    ToStr = s
End Function

