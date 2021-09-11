VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Fields
Public m_wrapper As Object
Public m_handle As IntPtr

Private Sub Class_Initialize()
  '
End Sub

Private Sub Class_Terminate()
  Set m_wrapper = Nothing
  Set m_handle = Nothing
End Sub

' Methods
Public Sub NewC(ByVal aWrapper As Object, ByVal aHandle As IntPtr)
  Set m_wrapper = aWrapper
  Set m_handle = aHandle
End Sub
Public Function New_HandleRef(ByVal aWrapper As Object, ByVal aHandle As IntPtr) As HandleRef
  Set New_HandleRef = New HandleRef
  Call New_HandleRef.NewC(aWrapper, aHandle)
End Function


'Public Shared Narrowing Operator CType(ByVal value As HandleRef) As IntPtr
Public Function op_CType(ByVal value As HandleRef) As IntPtr
  '
End Function

'Public Shared Function ToIntPtr(ByVal value As HandleRef) As IntPtr
Public Function ToIntPtr(ByVal value As HandleRef) As IntPtr
  Set ToIntPtr = New_IntPtr(value.m_handle.value)
End Function

' Properties
'Public ReadOnly Property Handle As IntPtr
Public Property Get Handle() As IntPtr
  Set Handle = m_handle
End Property

'Public ReadOnly Property Wrapper As Object
Public Property Get Wrapper() As Object
  Set Wrapper = m_wrapper
End Property

