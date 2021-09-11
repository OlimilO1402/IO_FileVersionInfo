Attribute VB_Name = "ModShift"
Option Explicit

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, Optional ByVal dwLen As Long = 4)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal module As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal addr As Long) As Long

Public Enum VirtualFreeTypes
  MEM_DECOMMIT = &H4000
  MEM_RELEASE = &H8000
End Enum
Public Enum VirtualAllocTypes
  MEM_COMMIT = &H1000
  MEM_RESERVE = &H2000
  MEM_RESET = &H8000
  MEM_LARGE_PAGES = &H20000000
  MEM_PHYSICAL = &H100000
  MEM_WRITE_WATCH = &H200000
End Enum
Public Enum VirtualAllocPageFlags
  PAGE_EXECUTE = &H10
  PAGE_EXECUTE_READ = &H20
  PAGE_EXECUTE_READWRITE = &H40
  PAGE_EXECUTE_WRITECOPY = &H80
  PAGE_NOACCESS = &H1
  PAGE_READONLY = &H2
  PAGE_READWRITE = &H4
  PAGE_WRITECOPY = &H8
  PAGE_GUARD = &H100
  PAGE_NOCACHE = &H200
  PAGE_WRITECOMBINE = &H400
End Enum
Private Type Memory
  address  As Long
  bytes    As Long
End Type

Private Const IDE_ADDROF_REL    As Long = 22

' by Donald, donald@xbeat.net
Private Const SHLCode As String = "8A4C240833C0F6C1E075068B442404D3E0C20800"
Private Const SHRCode As String = "8A4C240833C0F6C1E075068B442404D3E8C20800"
Private Const SARCode As String = "8A4C240833C0F6C1E075068B442404D3F8C20800"

Private m_memLHS                As Memory
Private m_memRHS                As Memory
Private m_memSAR                As Memory

Private m_clsHookLHS            As clsHook
Private m_clsHookRHS            As clsHook
Private m_clsHookSAR            As clsHook

Private m_blnInited             As Boolean

Public Property Get FastShiftingActive() As Boolean
    FastShiftingActive = m_blnInited
End Property

Public Sub InitFastShift()
  If Not m_blnInited Then
    m_memLHS = AsmToMem(SHLCode)
    m_memRHS = AsmToMem(SHRCode)
    m_memSAR = AsmToMem(SARCode)
    Set m_clsHookLHS = New clsHook
    Set m_clsHookRHS = New clsHook
    Set m_clsHookSAR = New clsHook
    m_clsHookLHS.Hook GetFncPtr(AddressOf ShiftLeft), m_memLHS.address
    m_clsHookRHS.Hook GetFncPtr(AddressOf ShiftRight), m_memRHS.address
    m_clsHookSAR.Hook GetFncPtr(AddressOf ShiftRightZ), m_memSAR.address
    m_blnInited = True
  End If
End Sub

Public Sub TermFastShift()
    If m_blnInited Then
    
        m_clsHookLHS.Unhook
        m_clsHookRHS.Unhook
        m_clsHookSAR.Unhook
        
        FreeMemory m_memLHS
        FreeMemory m_memRHS
        FreeMemory m_memSAR
        
        Set m_clsHookLHS = Nothing
        Set m_clsHookRHS = Nothing
        Set m_clsHookSAR = Nothing
        
        m_blnInited = False

    End If
End Sub

'Assembler Hex String in ausführbaren Speicher kopieren
Private Function AsmToMem(ByVal strAsm As String) As Memory
  Dim btAsm() As Byte
  Dim i       As Long
  Dim udtMem  As Memory
  ReDim btAsm(Len(strAsm) \ 2 - 1)
  For i = 0 To Len(strAsm) \ 2 - 1
    btAsm(i) = CByte("&H" & Mid$(strAsm, i * 2 + 1, 2))
  Next
  udtMem = AllocMemory(UBound(btAsm) + 1, , PAGE_EXECUTE_READWRITE)
  With udtMem
    Call CpyMem(ByVal .address, btAsm(0), UBound(btAsm) + 1)
    Call VirtualProtect(.address, .bytes, PAGE_EXECUTE_READ, 0)
  End With
  AsmToMem = udtMem
End Function
Private Function GetFncPtr(ByVal addrof As Long) As Long
  Dim pAddr As Long
  If IsRunningInIDE_DirtyTrick() Then
    'Wird das Programm aus der Entwicklungsumgebung heraus ausgeführt,
    'befindet sich der eigentliche Funktionszeiger bei (AddressOf X) + 22,
    'AddressOf X selber zeigt nur auf einen Stub. (getestet mit VB 6)
    Call CpyMem(pAddr, ByVal addrof + IDE_ADDROF_REL, 4)
    If IsBadCodePtr(pAddr) Then pAddr = addrof
  Else
      pAddr = addrof
  End If
  GetFncPtr = pAddr
End Function

Private Function AllocMemory(ByVal bytes As Long, Optional ByVal lpAddr As Long = 0, Optional ByVal PageFlags As VirtualAllocPageFlags = PAGE_READWRITE) As Memory
  With AllocMemory
    .address = VirtualAlloc(lpAddr, bytes, MEM_COMMIT, PageFlags)
    .bytes = bytes
  End With
End Function

Private Function FreeMemory(udtMem As Memory) As Boolean
  Call VirtualFree(udtMem.address, udtMem.bytes, MEM_DECOMMIT)
  udtMem.address = 0
  udtMem.bytes = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Public Function ShiftLeft(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Dim mask As Long
  
    Select Case ShiftCount
    
        Case 1 To 31
            ' mask out bits that are pushed over the edge anyway
            mask = Pow2(31 - ShiftCount)
            ShiftLeft = value And (mask - 1)
            ' shift
            ShiftLeft = ShiftLeft * Pow2(ShiftCount)
            ' set sign bit
            If value And mask Then
                ShiftLeft = ShiftLeft Or &H80000000
            End If
            
        Case 0
            ' ret unchanged
            ShiftLeft = value
            
    End Select
End Function

Public Function ShiftRightZ(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Select Case ShiftCount
    
        Case 1 To 31
            If value And &H80000000 Then
                ShiftRightZ = (value And Not &H80000000) \ 2
                ShiftRightZ = ShiftRightZ Or &H40000000
                ShiftRightZ = ShiftRightZ \ Pow2(ShiftCount - 1)
            Else
                ShiftRightZ = value \ Pow2(ShiftCount)
            End If
            
        Case 0
            ' ret unchanged
            ShiftRightZ = value
            
    End Select
End Function

Public Static Function ShiftRight(ByVal value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20011009
    Dim lPow2(0 To 30) As Long
    Dim i As Long
    
    Select Case ShiftCount
        Case 0
            ShiftRight = value
            
        Case 1 To 30
            If i = 0 Then
                lPow2(0) = 1
                For i = 1 To 30
                    lPow2(i) = 2 * lPow2(i - 1)
                Next
            End If
            
            If value And &H80000000 Then
                ShiftRight = value \ lPow2(ShiftCount)
                If ShiftRight * lPow2(ShiftCount) <> value Then
                    ShiftRight = ShiftRight - 1
                End If
            Else
                ShiftRight = value \ lPow2(ShiftCount)
            End If
            
        Case 31
            If value And &H80000000 Then
                ShiftRight = -1
            Else
                ShiftRight = 0
            End If
            
    End Select
End Function

Public Static Function Pow2(ByVal Exponent As Long) As Long
    ' by Donald, donald@xbeat.net, 20001217
    Dim alPow2(0 To 31) As Long
    Dim i As Long
  
    Select Case Exponent
    
        Case 0 To 31
            ' initialize lookup table
            If alPow2(0) = 0 Then
                alPow2(0) = 1
                For i = 1 To 30
                    alPow2(i) = alPow2(i - 1) * 2
                Next
                alPow2(31) = &H80000000
            End If
            
            ' return
            Pow2 = alPow2(Exponent)
            
    End Select
End Function

' http://www.activevb.de/tipps/vb6tipps/tipp0347.html
Private Function IsRunningInIDE_DirtyTrick() As Boolean
  On Error GoTo NotCompiled
  
    Debug.Print 1 / 0
    Exit Function
    
NotCompiled:
    IsRunningInIDE_DirtyTrick = True
    Exit Function
End Function

