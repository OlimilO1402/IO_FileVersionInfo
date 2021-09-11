Attribute VB_Name = "ModConstructors"
Option Explicit

Public Function New_StringBuilder(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
  Set New_StringBuilder = New StringBuilder
  Call New_StringBuilder.NewC(Value, startIndex, Length, Capacity, maxCapacity)
End Function

Public Function New_IntPtr(LngVal As Long) As IntPtr
  Set New_IntPtr = New IntPtr
  Call New_IntPtr.NewC(LngVal)
End Function

Public Function New_HandleRef(ByVal aWrapper As Object, ByVal aHandle As IntPtr) As HandleRef
  Set New_HandleRef = New HandleRef
  Call New_HandleRef.NewC(aWrapper, aHandle)
End Function

'Public Function New_Hook(ByVal pAddr As Long, ByVal NewAddr As Long, Optional ProxyAddr As Long) As clsHook
'  Set New_Hook = New Hook
'  Call New_Hook.NewC(pAddr, NewAddr, ProxyAddr)
'End Function

