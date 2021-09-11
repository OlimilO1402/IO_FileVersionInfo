Attribute VB_Name = "Win32Native"
Option Explicit
'Lutz Röders .net Reflector
'mscorlib.CommonLanguageRuntimeLibrary.Win32Native
'Namespace
'Microsoft.Win32
Public Const NNULL As Long = 0
#If defUnicode Then
  Private Declare Function plstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  'Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
  Private Declare Function plstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal dst As Long, ByVal src As Long) As Long
#Else
  Private Declare Function plstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
  'Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
  Private Declare Function plstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal dst As Long, ByVal src As Long) As Long
#End If
Private Declare Function plstrlenA Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function plstrlenW Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function plstrcpyA Lib "kernel32.dll" Alias "lstrcpyA" (ByVal dst As Long, ByVal src As Long) As Long
Private Declare Function plstrcpyW Lib "kernel32.dll" Alias "lstrcpyW" (ByVal dst As Long, ByVal src As Long) As Long

Public Function lstrcpy(ByVal dst As StringBuilder, ByVal src As IntPtr) As IntPtr
  'Call dst.NewC(, , , plstrlenW(src))
  Set lstrcpy = New_IntPtr(plstrcpy(dst.SPtr, src.value))
End Function

Public Function lstrlen(ByVal pStrVal As Long) As Long
  lstrlen = plstrlenA(pStrVal)
End Function

