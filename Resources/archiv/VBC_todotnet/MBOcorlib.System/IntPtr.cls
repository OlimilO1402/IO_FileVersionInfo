VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Structure IntPtr
'     Member von: System
'Zusammenfassung:
' Ein plattformabhängiger Typ zur Darstellung von Zeigern und Handles.

'Private mObj As Object
Private mObj As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (ByRef Destination As Any, ByVal Length As Long)

'Private Sub Class_Terminate()
'  Call Free
'End Sub

'eigentlich egal was hier übergeben wird, solange es Int32 ist.
'wenn jedoch das interne Objekt mObj gebraucht wird, dann hier
'immer den ObjPtr übergeben!
Public Property Let value(ObjPtrVal As Long)
  'Call RtlMoveMemory(ByVal VarPtr(mObj), ObjPtrVal, 4)
  mObj = ObjPtrVal
End Property
Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
Attribute value.VB_MemberFlags = "40"
  'Value = ObjPtr(mObj)
  value = mObj
  'value = VarPtr(mObj)
End Property

'hier mal anders, den VarPtr übergeben
Public Sub NewC(VarPtrObj As Long)
  'Call RtlMoveMemory(ByVal VarPtr(mObj), ByVal VarPtrObj, 4)
  mObj = VarPtrObj
End Sub
'in ein Modul ModConstructors
Public Function New_IntPtr(LngVal As Long) As IntPtr
  Set New_IntPtr = New IntPtr
  Call New_IntPtr.NewC(LngVal)
End Function

Public Property Get VPtr() As Long
  'liefert den VarPtr des internen Long, um darauf schreiben zu können
  VPtr = VarPtr(mObj)
End Property
Public Function VpOpToString(Optional StrName As String) As String
  If Len(StrName) = 0 Then StrName = "mObj"
  VpOpToString = "VarPtr(" & StrName & "): " & CStr(VarPtr(mObj)) & ";     " & "ObjPtr(" & StrName & "): " & ToString
End Function

'nicht von .net:
Public Property Get Obj() As Object
  'Set Obj = mObj
End Property

Public Sub Free()
  Call RtlZeroMemory(ByVal VarPtr(mObj), 4)
End Sub

'##########
'Public Overrides Function Equals(ByVal obj As Object) As Boolean
Public Function Equals(ByVal Obj As Object) As Boolean
'Zusammenfassung:
' Gibt einen Wert zurück, der angibt, ob diese Instanz gleich einem angegebenen Objekt ist.
'Parameter:
' obj: Ein Objekt, das mit dieser Instanz verglichen werden soll, oder null.
'Rückgabewerte:
' true, wenn obj eine Instanz von System.IntPtr ist, deren Wert gleich dem Wert dieser Instanz ist, andernfalls false.
  'If ObjPtr(Obj) = ObjPtr(mObj) Then Equals = True
  Equals = (ObjPtr(Obj) = mObj)
End Function

'Public Shared ReadOnly Property Size() As Integer
Public Property Get Size() As Long
'Zusammenfassung:
' Ruft die Größe dieser Instanz ab.
' Die Größe eines Zeigers bzw. Handles auf dieser Plattform in Bytes.
' Der Wert dieser Eigenschaft
' auf einer 32-Bit-Plattform ist 4 und
' auf einer 64-Bit-Plattform 8.
  Size = LenB(mObj)
End Property

'Public Function ToInt32() As Integer
Public Function ToInt32() As Long
'Zusammenfassung:
' Konvertiert den Wert dieser Instanz in eine 32-Bit-Ganzzahl mit Vorzeichen.
'Rückgabewerte:
' Eine 32-Bit-Ganzzahl mit Vorzeichen, die gleich dem Wert dieser Instanz ist.
  ToInt32 = mObj 'ObjPtr(mObj)
End Function

'Public Overrides Function ToString() As String
Public Function ToString() As String
'Zusammenfassung:
' Konvertiert den Wert dieser Instanz in die entsprechende Zeichenfolgenentsprechung.
'Rückgabewerte:
' Die Zeichenfolgenentsprechung des Wertes dieser Instanz.
  'ToString = CStr(ObjPtr(mObj))
  ToString = CStr(mObj)
End Function

'Public ReadOnly Shared Dim Zero As System.IntPtr
Public Property Get Zero() As IntPtr
'Zusammenfassung:
' Ein schreibgeschütztes Feld, das einen mit 0 initialisierten Zeiger oder ein Handle darstellt.
  Set Zero = New IntPtr
End Property

