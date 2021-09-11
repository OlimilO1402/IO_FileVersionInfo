VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Marshal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Fields
Private HIWORDMASK As IntPtr
'Private IID_IUnknown As Guid
Private Const LMEM_FIXED As Integer = 0
Private Const LMEM_MOVEABLE As Integer = 2
'Private Const s_strConvertedTypeInfoAssemblyDesc As String = "Type dynamically generated from ITypeInfo's"
'Private Const s_strConvertedTypeInfoAssemblyName As String = "InteropDynamicTypes"
'Private Const s_strConvertedTypeInfoAssemblyTitle As String = "Interop Dynamic Types"
'Private Const s_strConvertedTypeInfoNameSpace As String = "InteropDynamicTypes"
Public SystemDefaultCharSize As Long
Public SystemMaxDBCSCharSize As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal pDst As Long, ByVal pSrc As Long, ByVal Length As Long)

'Shared Sub New()
Private Sub Class_Initialize()
  Set HIWORDMASK = New_IntPtr(CLng(-65536))
  'Marshal.IID_IUnknown = New Guid("00000000-0000-0000-C000-000000000046")
  'Dim numArray1 As SByte() = New SByte(4  - 1) {}
  'ReDim numArray1(0 To 4 - 1) As Byte
  'numArray1(0) = 65
  'numArray1(1) = 65
  SystemDefaultCharSize = 1 '(3 - Win32Native.lstrlen(numArray1))
  'SystemMaxDBCSCharSize = Marshal.GetSystemMaxDBCSCharSize
End Sub




'Public Shared Function AddRef(ByVal pUnk As System.IntPtr) As Integer

'Public Shared Function AllocCoTaskMem(ByVal cb As Integer) As System.IntPtr

'Public Shared Function AllocHGlobal(ByVal cb As Integer) As System.IntPtr

'Public Shared Function AllocHGlobal(ByVal cb As System.IntPtr) As System.IntPtr

'Public Shared Function BindToMoniker(ByVal monikerName As String) As Object

'Public Shared Sub ChangeWrapperHandleStrength(ByVal otp As Object, ByVal fIsWeak As Boolean)

'Public Shared Sub Copy(ByVal source() As Byte, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Char, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Double, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Integer, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Long, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Short, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source() As Single, ByVal startIndex As Integer, ByVal destination As System.IntPtr, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Byte, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Char, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Double, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Integer, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Long, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Short, ByVal startIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal source As System.IntPtr, ByVal destination() As Single, ByVal startIndex As Integer, ByVal length As Integer)

'Public Shared Function CreateWrapperOfType(ByVal o As Object, ByVal t As System.Type) As Object

'Public Shared Sub DestroyStructure(ByVal ptr As System.IntPtr, ByVal structuretype As System.Type)

'Public Shared Sub FreeBSTR(ByVal ptr As System.IntPtr)

'Public Shared Sub FreeCoTaskMem(ByVal ptr As System.IntPtr)

'Public Shared Sub FreeHGlobal(ByVal hglobal As System.IntPtr)

'Public Shared Function GenerateGuidForType(ByVal type As System.Type) As System.Guid

'Public Shared Function GenerateProgIdForType(ByVal type As System.Type) As String

'Public Shared Function GetActiveObject(ByVal progID As String) As Object

'Public Shared Function GetComInterfaceForObject(ByVal o As Object, ByVal T As System.Type) As System.IntPtr

'Public Shared Function GetComObjectData(ByVal obj As Object, ByVal key As Object) As Object

'Public Shared Function GetComSlotForMethodInfo(ByVal m As System.Reflection.MemberInfo) As Integer

'Public Shared Function GetEndComSlot(ByVal t As System.Type) As Integer

'Public Shared Function GetExceptionCode() As Integer

'Public Shared Function GetExceptionPointers() As System.IntPtr

'Public Shared Function GetHINSTANCE(ByVal m As System.Reflection.Module) As System.IntPtr

'Public Shared Function GetHRForException(ByVal e As System.Exception) As Integer

'Public Shared Function GetHRForLastWin32Error() As Integer

'Public Shared Function GetIDispatchForObject(ByVal o As Object) As System.IntPtr

'Public Shared Function GetITypeInfoForType(ByVal t As System.Type) As System.IntPtr

'Public Shared Function GetIUnknownForObject(ByVal o As Object) As System.IntPtr

'Public Shared Function GetLastWin32Error() As Integer

'Public Shared Function GetManagedThunkForUnmanagedMethodPtr(ByVal pfnMethodToWrap As System.IntPtr, ByVal pbSignature As System.IntPtr, ByVal cbSignature As Integer) As System.IntPtr

'Public Shared Function GetMethodInfoForComSlot(ByVal t As System.Type, ByVal slot As Integer, ByRef memberType As System.Runtime.InteropServices.ComMemberType) As System.Reflection.MemberInfo

'Public Shared Sub GetNativeVariantForObject(ByVal obj As Object, ByVal pDstNativeVariant As System.IntPtr)

'Public Shared Function GetObjectForIUnknown(ByVal pUnk As System.IntPtr) As Object

'Public Shared Function GetObjectForNativeVariant(ByVal pSrcNativeVariant As System.IntPtr) As Object

'Public Shared Function GetObjectsForNativeVariants(ByVal aSrcNativeVariant As System.IntPtr, ByVal cVars As Integer) As Object()

'Public Shared Function GetStartComSlot(ByVal t As System.Type) As Integer

'Public Shared Function GetThreadFromFiberCookie(ByVal cookie As Integer) As System.Threading.Thread

'Public Shared Function GetTypedObjectForIUnknown(ByVal pUnk As System.IntPtr, ByVal t As System.Type) As Object
  
'Public Shared Function GetTypeForITypeInfo(ByVal piTypeInfo As System.IntPtr) As System.Type

'Public Shared Function GetTypeInfoName(ByVal pTI As System.Runtime.InteropServices.UCOMITypeInfo) As String

'Public Shared Function GetTypeLibGuid(ByVal pTLB As System.Runtime.InteropServices.UCOMITypeLib) As System.Guid

'Public Shared Function GetTypeLibGuidForAssembly(ByVal asm As System.Reflection.Assembly) As System.Guid

'Public Shared Function GetTypeLibLcid(ByVal pTLB As System.Runtime.InteropServices.UCOMITypeLib) As Integer

'Public Shared Function GetTypeLibName(ByVal pTLB As System.Runtime.InteropServices.UCOMITypeLib) As String

'Public Shared Function GetUnmanagedThunkForManagedMethodPtr(ByVal pfnMethodToWrap As System.IntPtr, ByVal pbSignature As System.IntPtr, ByVal cbSignature As Integer) As System.IntPtr

'Public Shared Function IsComObject(ByVal o As Object) As Boolean

'Public Shared Function IsTypeVisibleFromCom(ByVal t As System.Type) As Boolean

'Public Shared Function NumParamBytes(ByVal m As System.Reflection.MethodInfo) As Integer

'Public Shared Function OffsetOf(ByVal t As System.Type, ByVal fieldName As String) As System.IntPtr

'Public Shared Sub Prelink(ByVal m As System.Reflection.MethodInfo)

'Public Shared Sub PrelinkAll(ByVal c As System.Type)

'Public Shared Function PtrToStringAnsi(ByVal ptr As System.IntPtr) As String
'Public Shared Function PtrToStringAnsi(ByVal ptr As System.IntPtr, ByVal len As Integer) As String

'Public Shared Function PtrToStringAuto(ByVal ptr As System.IntPtr) As String
'Public Shared Function PtrToStringAuto(ByVal ptr As System.IntPtr, ByVal len As Integer) As String
'      If (Marshal.SystemDefaultCharSize <> 1) Then
'            Return Marshal.PtrToStringUni(ptr, len)
'      End If
'      Return Marshal.PtrToStringAnsi(ptr, len)
'End Function
Public Function PtrToStringAuto(ByVal Ptr As IntPtr, Optional ByVal sLen As Long) As String
  If (Win32Native.NNULL = Ptr) Then
    'Return Nothing
    Exit Function
    MsgBox "PtrToStringAuto: Ptr=Nullpointer"
  End If
'  If Marshal.IsWin32Atom(ptr) Then
'        Return Nothing
'  End If
  Dim num1 As Long: num1 = Win32Native.lstrlen(Ptr)
  Dim builder1 As StringBuilder: Set builder1 = New_StringBuilder(, , , num1 * 2)
  Call Win32Native.lstrcpy(builder1, Ptr)
  'Return builder1.ToString
  PtrToStringAuto = builder1.ToString
End Function

'Public Shared Function PtrToStringBSTR(ByVal ptr As System.IntPtr) As String

'Public Shared Function PtrToStringUni(ByVal ptr As System.IntPtr) As String

'Public Shared Function PtrToStringUni(ByVal ptr As System.IntPtr, ByVal len As Integer) As String

'Public Shared Sub PtrToStructure(ByVal ptr As System.IntPtr, ByVal structure As Object)
'Public Shared Function PtrToStructure(ByVal ptr As System.IntPtr, ByVal structureType As System.Type) As Object
Public Sub PtrToStructure(ByVal Ptr As IntPtr, ByVal structure As Object)
  'hier wird nicht der Pointer übertragen, sondern vielmehr der gesamte Speicherbereich
  'in das Objekt structure hineinkopiert
  'Statt structure.value könnte man auch schreiben:
  'ByVal (DeRef(ObjPtr(structure)) + 52)
  'könnte man auch das Objekt daraufzeigen lassen?
  'bzw den Pointer tatsächlich verbiegen?
  'nee ik gloob det jeht nit
  Call RtlMoveMemory(ByVal structure.value, ByVal Ptr.value, structure.ByteLength)
End Sub


'Public Shared Function QueryInterface(ByVal pUnk As System.IntPtr, ByRef iid As System.Guid, ByRef ppv As System.IntPtr) As Integer

'##############################'   Alle Read   '##############################'
'Public Shared Function ReadByte(ByVal ptr As Object, ByVal ofs As Integer) As Byte
'Public Shared Function ReadByte(ByVal ptr As System.IntPtr) As Byte
'Public Shared Function ReadByte(ByVal ptr As System.IntPtr, ByVal ofs As Integer) As Byte


'Public Shared Function ReadInt16(ByVal ptr As Object, ByVal ofs As Integer) As Short
'Public Shared Function ReadInt16(ByVal ptr As System.IntPtr) As Short
'Public Shared Function ReadInt16(ByVal ptr As System.IntPtr, ByVal ofs As Integer) As Short
Public Function ReadInt16(ByVal Ptr As IntPtr, Optional ByVal ofs As Long) As Integer
  '                  pDst             , pSrc
  Call RtlMoveMemory(VarPtr(ReadInt16), (Ptr + ofs), 2)
End Function

'Public Shared Function ReadInt32(ByVal ptr As Object, ByVal ofs As Integer) As Integer
'Public Shared Function ReadInt32(ByVal ptr As System.IntPtr) As Integer
'Public Shared Function ReadInt32(ByVal ptr As System.IntPtr, ByVal ofs As Integer) As Integer
Public Function ReadInt32(Ptr As Long, Optional ByVal ofs As Long) As Long
  Call RtlMoveMemory(VarPtr(ReadInt32), (Ptr + ofs), 4)
End Function

'Public Shared Function ReadInt64(ByVal ptr As Object, ByVal ofs As Integer) As Long
'Public Shared Function ReadInt64(ByVal ptr As System.IntPtr) As Long
'Public Shared Function ReadInt64(ByVal ptr As System.IntPtr, ByVal ofs As Integer) As Long

'Public Shared Function ReadIntPtr(ByVal ptr As Object, ByVal ofs As Integer) As System.IntPtr
'Public Shared Function ReadIntPtr(ByVal ptr As System.IntPtr) As System.IntPtr
'Public Shared Function ReadIntPtr(ByVal ptr As System.IntPtr, ByVal ofs As Integer) As System.IntPtr

'Public Shared Function ReAllocCoTaskMem(ByVal pv As System.IntPtr, ByVal cb As Integer) As System.IntPtr
'Public Shared Function ReAllocHGlobal(ByVal pv As System.IntPtr, ByVal cb As System.IntPtr) As System.IntPtr

'Public Shared Function Release(ByVal pUnk As System.IntPtr) As Integer
'Public Shared Function ReleaseComObject(ByVal o As Object) As Integer
'Public Shared Sub ReleaseThreadCache()
'Public Shared Function SetComObjectData(ByVal obj As Object, ByVal key As Object, ByVal data As Object) As Boolean
'Public Shared Function SizeOf(ByVal structure As Object) As Integer
'Public Shared Function SizeOf(ByVal t As System.Type) As Integer
'Public Shared Function StringToBSTR(ByVal s As String) As System.IntPtr
'Public Shared Function StringToCoTaskMemAnsi(ByVal s As String) As System.IntPtr
'Public Shared Function StringToCoTaskMemAuto(ByVal s As String) As System.IntPtr
'Public Shared Function StringToCoTaskMemUni(ByVal s As String) As System.IntPtr
'Public Shared Function StringToHGlobalAnsi(ByVal s As String) As System.IntPtr
'Public Shared Function StringToHGlobalAuto(ByVal s As String) As System.IntPtr
'Public Shared Function StringToHGlobalUni(ByVal s As String) As System.IntPtr
'Public Shared Sub StructureToPtr(ByVal structure As Object, ByVal ptr As System.IntPtr, ByVal fDeleteOld As Boolean)
'Public ReadOnly Shared Dim SystemDefaultCharSize As Integer
'Public ReadOnly Shared Dim SystemMaxDBCSCharSize As Integer
'Public Shared Sub ThrowExceptionForHR(ByVal errorCode As Integer)
'Public Shared Sub ThrowExceptionForHR(ByVal errorCode As Integer, ByVal errorInfo As System.IntPtr)
'Public Shared Function UnsafeAddrOfPinnedArrayElement(ByVal arr As System.Array, ByVal index As Integer) As System.IntPtr

'##############################'   Alle Read   '##############################'
'Public Shared Sub WriteByte(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As Byte)
'Public Shared Sub WriteByte(ByVal ptr As System.IntPtr, ByVal val As Byte)
'Public Shared Sub WriteByte(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As Byte)

'Public Shared Sub WriteInt16(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As Char)
'Public Shared Sub WriteInt16(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As Short)
'Public Shared Sub WriteInt16(ByVal ptr As System.IntPtr, ByVal val As Char)
'Public Shared Sub WriteInt16(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As Char)
'Public Shared Sub WriteInt16(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As Short)
'Public Shared Sub WriteInt16(ByVal ptr As System.IntPtr, ByVal val As Short)

'Public Shared Sub WriteInt32(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As Integer)
'Public Shared Sub WriteInt32(ByVal ptr As System.IntPtr, ByVal val As Integer)
'Public Shared Sub WriteInt32(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As Integer)
'Public Shared Sub WriteInt64(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As Long)
'Public Shared Sub WriteInt64(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As Long)
'Public Shared Sub WriteInt64(ByVal ptr As System.IntPtr, ByVal val As Long)
'Public Shared Sub WriteIntPtr(ByVal ptr As Object, ByVal ofs As Integer, ByVal val As System.IntPtr)
'Public Shared Sub WriteIntPtr(ByVal ptr As System.IntPtr, ByVal ofs As Integer, ByVal val As System.IntPtr)
'Public Shared Sub WriteIntPtr(ByVal ptr As System.IntPtr, ByVal val As System.IntPtr)

