Attribute VB_Name = "UnsafeNativeMethods"
Option Explicit
#If defUnicode Then
  Private Declare Function pGetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
  Private Declare Function pGetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
  Private Declare Function pVerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
  Private Declare Function pVerLanguageName Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
#Else
  Private Declare Function pGetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
  Private Declare Function pGetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
  Private Declare Function pVerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
  Private Declare Function pVerLanguageName Lib "kernel32.dll" Alias "VerLanguageNameW " (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
#End If
Private Declare Function pGetFileVersionInfoSizeW Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, ByRef lpdwHandle As Long) As Long
Private Declare Function pGetFileVersionInfoSizeA Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As Long, ByRef lpdwHandle As Long) As Long
Private Declare Function pGetFileVersionInfoW Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function pGetFileVersionInfoA Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function pVerQueryValueW Lib "version.dll" Alias "VerQueryValueW" (ByVal pBlock As Long, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Function pVerQueryValueA Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, ByRef puLen As Long) As Long
'wie kann die VerlanguageName-API in lpbuffer schreiben wenn es ByVal ist?
Private Declare Function pVerLanguageNameW Lib "kernel32.dll" Alias "VerLanguageNameW " (ByVal wLang As Long, ByVal szLang As Long, ByVal nSize As Long) As Long
Private Declare Function pVerLanguageNameA Lib "kernel32.dll" Alias "VerLanguageNameA " (ByVal wLang As Long, ByVal szLang As Long, ByVal nSize As Long) As Long
Private Declare Function pVerLanguageNameAA Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Declare Function pVerLanguageNameAAP Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As Long, ByVal nSize As Long) As Long

'Public Shared Function GetFileVersionInfo(ByVal lptstrFilename As String, ByVal dwHandle As Integer, ByVal dwLen As Integer, ByVal lpData As HandleRef) As Boolean
Public Function GetFileVersionInfo(ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByVal lpData As HandleRef) As Boolean
  GetFileVersionInfo = CBool(pGetFileVersionInfoA(lptstrFilename, dwHandle, dwLen, ByVal lpData.m_handle.value)) ' <> 0)
End Function

'Public Shared Function GetFileVersionInfoSize(ByVal lptstrFilename As String, <Out> ByRef handle As Integer) As Integer
Public Function GetFileVersionInfoSize(ByVal lptstrFilename As String, ByRef Handle As Long) As Long
  GetFileVersionInfoSize = pGetFileVersionInfoSize(lptstrFilename, Handle)
End Function

'Public Function VerQueryValue(ByVal pBlock As HandleRef, ByVal lpSubBlock As String, <[In], Out> ByRef lplpBuffer As IntPtr, <Out> ByRef len As Long) As Boolean
Public Function VerQueryValue(ByVal pBlock As HandleRef, ByVal lpSubBlock As String, ByRef lplpBuffer As IntPtr, ByRef llen As Long) As Boolean
  VerQueryValue = CBool(pVerQueryValueA(pBlock.m_handle.value, lpSubBlock, ByVal lplpBuffer.VPtr, llen)) ' <> 0)
End Function

'Public Shared Function VerLanguageName(ByVal langID As Integer, ByVal lpBuffer As StringBuilder, ByVal nSize As Integer) As Integer
Public Function VerLanguageName(ByVal langID As Long, ByVal lpBuffer As StringBuilder, ByVal nSize As Long) As Long
  Dim StrRet As String '* 256
  StrRet = String$(256, vbNullChar)
  VerLanguageName = pVerLanguageNameAA(langID, StrRet, nSize)
  StrRet = Left$(StrRet, VerLanguageName)
  Call lpBuffer.Append(StrRet)
  
'  VerLanguageName = pVerLanguageNameAAP(langID, lpBuffer.SPtr, nSize)
'  lpBuffer.Length = VerLanguageName
  'wird die Länge des zurückgegebenen Strings irgendwie hier übermittelt?
  'wenn ja, dann kann sie in den lpbuffer geschrieben werden
End Function
