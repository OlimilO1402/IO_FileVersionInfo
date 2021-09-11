VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
' function redirection class
'
' [rm_code]
'***********************************
' Thanks to:
' EBArtSoft's API HOOK Demo II
'***********************************

Private Declare Function VirtualProtect Lib "kernel32" ( _
    lpAddress As Any, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    lpflOldProtect As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long _
)

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Private Declare Function GetModuleHandle Lib "kernel32" _
Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String _
) As Long

Private Const PAGE_EXECUTE_READWRITE    As Long = &H40&

Private blnHooked       As Boolean
Private lpOldAddr       As Long

Private btOldASM(4)     As Byte
Private btReal(31)      As Byte

Public Function Unhook() As Boolean
    If blnHooked Then
        blnHooked = PutMem(lpOldAddr, VarPtr(btOldASM(0)), UBound(btOldASM) + 1)
        Unhook = blnHooked
        blnHooked = False
    End If
End Function

Public Property Get OldAddress() As Long
    OldAddress = lpOldAddr
End Property

Public Property Get Hooked() As Boolean
    Hooked = blnHooked
End Property

Public Function Hook( _
    ByVal pAddr As Long, _
    ByVal NewAddr As Long, _
    Optional ProxyAddr As Long _
) As Boolean

    Dim hModule     As Long
    Dim hFnc        As Long

    If Not blnHooked Then
        lpOldAddr = pAddr
        hFnc = pAddr
    
        If Not GetMem(hFnc, VarPtr(btOldASM(0)), UBound(btOldASM) + 1) Then
            Exit Function
        End If
    
        If ProxyAddr <> 0 Then
            CopyMemory btReal(0), btOldASM(0), UBound(btOldASM) + 1
            Redirect VarPtr(btReal(UBound(btOldASM) + 1)), lpOldAddr + UBound(btOldASM) + 1
            Redirect ProxyAddr, VarPtr(btReal(0))
        End If
    
        blnHooked = Redirect(hFnc, NewAddr)
        Hook = blnHooked
    End If
End Function

Public Function Redirect(ByVal OldAddr As Long, ByVal NewAddr As Long) As Boolean
    Dim btAsm(4)    As Byte
    Dim lngNewAddr  As Long

    lngNewAddr = NewAddr - OldAddr - (UBound(btAsm) + 1)

    btAsm(0) = &HE9
    CopyMemory btAsm(1), lngNewAddr, 4

    Redirect = PutMem(OldAddr, VarPtr(btAsm(0)), UBound(btAsm) + 1)

End Function

Private Function GetMem( _
    ByVal lpAddr As Long, _
    ByVal pData As Long, _
    ByVal dlen As Long _
) As Boolean

    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal pData, ByVal lpAddr, dlen
    
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    GetMem = True
End Function

Private Function PutMem( _
    ByVal lpAddr As Long, _
    ByVal pData As Long, _
    ByVal dlen As Long _
) As Boolean

    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal lpAddr, ByVal pData, dlen
    
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    PutMem = True
End Function


