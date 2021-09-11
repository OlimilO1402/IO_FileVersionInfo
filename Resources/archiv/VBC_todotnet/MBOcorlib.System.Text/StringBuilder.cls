VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class StringBuilder
'          Inherits System.Object
'     Member von: System.Text
'Zusammenfassung:
' Stellt eine veränderbare Zeichenfolge dar. Diese Klasse kann nicht geerbt werden.
'
'Comments:
'in .NET: diese Klasse hält den zu verknüpfenden String selbst,
'Deshalb nicht GlobalMultiUse sondern nur Instancing = MultiUse
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr As Long, ByVal BLen As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal olestr As Long, ByVal BLen As Long) As String
Private mStrValue As String  'm_sString
Private mCachedStr As String
Private mLength As Long   'in Byte!!!! mLength = 2 * Len(String) = LenB(String)
Private mCapacity As Long 'in Byte!! ist eigentlich LenB(mStrValue)
Private mMaxCapacity As Long '= 2147483647 'braucht man das überhaupt?
'in .NET 16, in Mono 8192
Private Const constDefaultCapacity = 16 '8192 '16384 '16 'm_iChunkSize
'Private Const constDefaultCapacity = 8192
'Many Thanks To:
'* KonRad L.M. Rudolph  www.ActiveVB.de
'* Klaus Langbein       www.ActiveVB.de
'* Matthew Curland for "Advanced Visual Basic 6" www.powervb.com
'* Steve McMahon (steve@vbaccelerator.com)
'* Jost Schwider www.vbtec.de
'* Donald www.xbeat.net -> vbspeed
'* Mono, all the guys around it and Novell,
'  especially the authors of System.String.cs, and System.StringBuilder.cs
'* #develop

Private Sub Class_Initialize()
  mMaxCapacity = 2147483647
  ' The default allocation: 8192 characters.
  'mCapacity = constDefaultCapacity
  mStrValue = InternalAllocateStr(constDefaultCapacity * 2)
End Sub

'Public Sub New()
'Public Sub New(ByVal capacity As Integer)
'Public Sub New(ByVal capacity As Integer, ByVal maxCapacity As Integer)
'Public Sub New(ByVal value As String)
'Public Sub New(ByVal value As String, ByVal capacity As Integer)
'Public Sub New(ByVal value As String, ByVal startIndex As Integer, ByVal length As Integer, ByVal capacity As Integer)
'    public StringBuilder(string value, int startIndex, int length, int capacity)
'    {
'      // first, check the parameters and throw appropriate exceptions if needed
'      if (null == value)
'        value = "";
'
'      // make sure startIndex is zero or positive
'      if (startIndex < 0)
'        throw new System.ArgumentOutOfRangeException ("startIndex", startIndex, "StartIndex cannot be less than zero.");
'
'      // make sure length is zero or positive
'      if(length < 0)
'        throw new System.ArgumentOutOfRangeException ("length", length, "Length cannot be less than zero.");
'
'      if (capacity < 0)
'        throw new System.ArgumentOutOfRangeException ("capacity", capacity, "capacity must be greater than zero.");
'
'      // make sure startIndex and length give a valid substring of value
'      // re-ordered to avoid possible integer overflow
'      if (startIndex > value.Length - length)
'        throw new System.ArgumentOutOfRangeException ("startIndex", startIndex, "StartIndex and length must refer to a location within the string.");
'
'      if (capacity == 0)
'        capacity = constDefaultCapacity;
'
'      _str = String.InternalAllocateStr ((length > capacity) ? length : capacity);
'      if (length > 0)
'        String.InternalStrcpy(_str, 0, value, startIndex, length);
'
'      _length = length;
'    }
Public Sub NewC(Optional ByVal sValue As String, Optional ByVal startIndex As Long, Optional ByVal lLength As Long, Optional ByVal aCapacity As Long, Optional ByVal maxCapacity As Long)
Dim n As Long
  If (aCapacity < 0) Then
    'Call Throw New_ArgumentOutOfRangeException("capacity", CStr(capacity), "capacity cannot be less than zero.")
    MsgBox "ArgumentOutOfRangeException :" & "capacity : " & CStr(aCapacity) & " capacity cannot be less than zero."
    Exit Sub
  End If
  If (lLength < 0) Then
    'Call Throw New_ArgumentOutOfRangeException("length", CStr(length), "length cannot be less than zero.")
    MsgBox "ArgumentOutOfRangeException :" & "length : " & CStr(lLength) & " length cannot be less than zero."
    Exit Sub
  End If
  If (startIndex < 0) Then
    'Call Throw New_ArgumentOutOfRangeException("startIndex", CStr(startIndex), "StartIndex cannot be less than zero.")
    MsgBox "ArgumentOutOfRangeException :" & "startIndex: " & CStr(startIndex) & " StartIndex cannot be less than zero."
    Exit Sub
  End If
  'If sValue = vbNullString Then sValue = ""
  'Make sure Startindex and length give a valid substring of value
  're-ordered to avoid possible integer overflow
  n = Len(sValue)
  If (startIndex > (n - lLength)) Then
    'Call Throw New_ArgumentOutOfRangeException("startIndex", CStr(startIndex), "StartIndex and length must refer to a location within the string.")
    MsgBox "ArgumentOutOfRangeException :" & "startIndex: " & CStr(startIndex) & " StartIndex and length must refer to a location within the string."
    Exit Sub
  End If
  If n = 0 Then
    If (maxCapacity < 1) Then
      'Throw New_ArgumentOutOfRangeException("maxCapacity", Environment.GetResourceString("ArgumentOutOfRange_SmallMaxCapacity"))
      'MsgBox "ArgumentOutOfRangeException :" & "maxCapacity : " & CStr(maxCapacity) & " maxCapacity cannot be less than one."
      'Exit Sub
    End If
    If (aCapacity > maxCapacity) Then
      'Throw New_ArgumentOutOfRangeException("capacity", Environment.GetResourceString("ArgumentOutOfRange_Capacity"))
      'MsgBox "ArgumentOutOfRangeException :" & "capacity : " & CStr(aCapacity) & "   " & "maxCapacity : " & CStr(maxCapacity) & " capacity cannot be greater than maxCapacity."
      'Exit Sub
    End If
  End If
  If aCapacity = 0 Then
    aCapacity = constDefaultCapacity
  End If
  If maxCapacity > 0 Then mMaxCapacity = maxCapacity
  If lLength = 0 Then lLength = n
  Do While (aCapacity < lLength)
    aCapacity = (aCapacity * 2)
    If (aCapacity < 0) Then
      aCapacity = lLength
      Exit Do
    End If
  Loop
  'If n > mCapacity Then IAS = n Else IAS = mCapacity
  mStrValue = InternalAllocateStr(aCapacity * 2)
  If lLength > 0 Then
    'Call InternalStrcpy(mStrValue, 0, sValue, startIndex, lLength)
    MidB$(mStrValue, mLength + 1) = MidB$(sValue, startIndex * 2 + 1, lLength * 2)
    'Call Append(Mid$(sValue, startIndex + 1, lLength))
    mLength = lLength * 2
  End If
End Sub
'Put this in a standard Module:
'Public Function New_StringBuilder(Optional ByVal value As String, Optional ByVal startIndex As Long, Optional ByVal length As Long, Optional ByVal capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
'  Set New_StringBuilder = New StringBuilder
'  Call New_StringBuilder.NewC(value, startIndex, length, capacity, maxCapacity)
'End Function


'Public Function Append(ByVal value As Boolean) As System.Text.StringBuilder
'Public Function Append(ByVal value As Byte) As System.Text.StringBuilder
'Public Function Append(ByVal value() As Char) As System.Text.StringBuilder
'Public Function Append(ByVal value() As Char, ByVal startIndex As Integer, ByVal charCount As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Char) As System.Text.StringBuilder
'Public Function Append(ByVal value As Char, ByVal repeatCount As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Decimal) As System.Text.StringBuilder
'Public Function Append(ByVal value As Double) As System.Text.StringBuilder
'Public Function Append(ByVal value As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As Long) As System.Text.StringBuilder
'Public Function Append(ByVal value As Object) As System.Text.StringBuilder
'Public Function Append(ByVal value As Short) As System.Text.StringBuilder
'Public Function Append(ByVal value As Single) As System.Text.StringBuilder
'Public Function Append(ByVal value As String) As System.Text.StringBuilder
'Public Function Append(ByVal value As String, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.SByte) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt16) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt32) As System.Text.StringBuilder
'Public Function Append(ByVal value As System.UInt64) As System.Text.StringBuilder
'Public Function Append(ByVal varValue As Variant) As StringBuilder
'  mValue = mValue & CStr(varValue)
'  Set Append = Me
'End Function
Public Function Append(ByRef StrValue As String) As StringBuilder
Dim n As Long, lLen As Long, lLength As Long
Dim aCapacity As Long
  lLen = LenB(mStrValue)
  aCapacity = lLen
  n = LenB(StrValue)
  lLength = mLength + n 'die resultierende Länge
  'gut ist, wenn aCapacity immer größer ist!
  If aCapacity < lLength Then
    Do While (aCapacity < lLength)
      'mStrValue = mStrValue & Space$(mCapacity)
      'mStrValue = mStrValue & SysAllocStringLen(0&, mCapacity \ 2)
      
      aCapacity = (aCapacity * 2) '+ mCapacity 'die kapazität verdoppeln
    Loop
    mStrValue = mStrValue & InternalAllocateStr(aCapacity - lLen)
  End If
  'CopyMemory ByVal UnsignedAdd(StrPtr(mStrValue), mLength), ByVal StrPtr(StrValue), lLen
  MidB$(mStrValue, mLength + 1) = StrValue
  mLength = mLength + n
  Set Append = Me
End Function
Public Sub AppendS(ByRef StrValue As String) 'As StringBuilder
Dim n As Long, lLen As Long, lLength As Long
Dim aCapacity As Long
  lLen = LenB(mStrValue)
  aCapacity = lLen
  n = LenB(StrValue)
  lLength = mLength + n 'die resultierende Länge
  'gut ist, wenn aCapacity immer größer ist!
  If aCapacity < lLength Then
    Do While (aCapacity < lLength)
      'mStrValue = mStrValue & Space$(mCapacity)
      'mStrValue = mStrValue & SysAllocStringLen(0&, mCapacity \ 2)
      
      aCapacity = (aCapacity * 2) '+ mCapacity 'die kapazität verdoppeln
    Loop
    mStrValue = mStrValue & InternalAllocateStr(aCapacity - lLen)
  End If
  'CopyMemory ByVal UnsignedAdd(StrPtr(mStrValue), mLength), ByVal StrPtr(StrValue), lLen
  MidB$(mStrValue, mLength + 1) = StrValue
  mLength = mLength + n
End Sub

'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal format As String, ByVal ParamArray args() As Object) As System.Text.StringBuilder
'Public Function AppendFormat(ByVal provider As System.IFormatProvider, ByVal format As String, ByVal ParamArray args() As Object) As System.Text.StringBuilder
Public Function AppendFormat(ByVal sFormat As String, ParamArray args()) As StringBuilder
  'Throw New_NotImplementedException("Not yet implemented")
  Set AppendFormat = Me
End Function

Public Function AppendLine(Optional StrVal As String) As StringBuilder
  If Len(StrVal) > 0 Then Call Append(StrVal)
  Set AppendLine = Append(vbNewLine)
End Function


Public Property Get capacity() As Long
  capacity = Len(mStrValue) 'mCapacity
End Property
Public Property Let capacity(LngVal As Long)
  'mCapacity = LngVal
  Call EnsureCapacity(LngVal)
End Property

'Public Function EnsureCapacity(ByVal capacity As Integer) As Integer
Public Function EnsureCapacity(ByVal reqCapacity As Long) As Long 'cInteger
Dim aCapacity As Long, lLength As Long
  aCapacity = LenB(mStrValue)
  lLength = reqCapacity * 2
  'If n >= aCapacity Then
  If aCapacity >= lLength Then
    EnsureCapacity = aCapacity
  Else
    'halt hier nicht so:
    'Do While (aCapacity < lLength)
    '  aCapacity = (aCapacity * 2)
    'Loop
    mStrValue = mStrValue & InternalAllocateStr(lLength - aCapacity)
    EnsureCapacity = aCapacity
  End If
End Function

'Public Function Equals(ByVal sb As System.Text.StringBuilder) As Boolean
Public Function Equals(ByVal SB As StringBuilder) As Boolean
  'Set Equals = New cBoolean
  If SB Is Me Then
    Equals = True
  Else
    If SB.ToString = mStrValue Then Equals = True
  End If
End Function

'Public Function Insert(ByVal index As Integer, ByVal value As Boolean) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Byte) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Char) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Decimal) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Double) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Integer) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Long) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Object) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Short) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As Single) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.SByte) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt16) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt32) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As System.UInt64) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As String) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value As String, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value() As Char) As System.Text.StringBuilder
'Public Function Insert(ByVal index As Integer, ByVal value() As Char, ByVal startIndex As Integer, ByVal charCount As Integer) As System.Text.StringBuilder
Public Function Insert(ByVal Index As Long, ByRef value As Variant, Optional ByVal startIndex As Long = 1, Optional ByVal charCount As Long = 1) As StringBuilder
'Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
Dim aCapacity As Long: aCapacity = LenB(mStrValue)
Dim lLen As Long
Dim lPos As Long
Dim lSize As Long
Dim sValue As String
  ' is iIndex within bounds?
  If (Index * 2 > mLength) Then
    Err.Raise 9
  Else
    sValue = CStr(value)
    lLen = LenB(sValue)
    If (mLength + lLen) > aCapacity Then
      aCapacity = aCapacity + aCapacity
      'aCapacity = aCapacity + lLen '+ mLength
      mStrValue = mStrValue & InternalAllocateStr(aCapacity - mLength)
    End If
    
    ' Move existing characters from current position
    lPos = UnsignedAdd(StrPtr(mStrValue), Index * 2)
    lSize = mLength - (Index * 2) '+ 4
    
    ' moving from iIndex to iIndex + lLen
    CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
    
    ' Insert new characters:
    CopyMemory ByVal lPos, ByVal StrPtr(sValue), lLen
    
    mLength = mLength + lLen
  End If
  Set Insert = Me
End Function

'Public Function Remove(ByVal startIndex As Integer, ByVal length As Integer) As System.Text.StringBuilder
Public Function Remove(ByVal startIndex As Long, ByVal length As Long) As StringBuilder
'Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
Dim lSrc As Long
Dim lDst As Long
Dim lSize As Long
  ' is iIndex within bounds?
  If (startIndex * 2 > mLength) Then
    'Err.Raise 9
  Else
    ' is there sufficient length?
    If ((startIndex + length) * 2 > mLength) Then
      'Err.Raise 9
    Else
      ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
      lSrc = UnsignedAdd(StrPtr(mStrValue), (startIndex + length) * 2)
      lDst = UnsignedAdd(StrPtr(mStrValue), startIndex * 2)
      lSize = (mLength - (startIndex + length) * 2)
      CopyMemory ByVal lDst, ByVal lSrc, lSize
      mLength = mLength - length * 2
    End If
  End If
  Set Remove = Me
'End Sub
End Function

Private Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
Dim lInstr As Long
  'Set Find = New cInteger
  If (lStartIndex > 0) Then
    lInstr = InStr(lStartIndex, mStrValue, sToFind, compare)
  Else
    lInstr = InStr(mStrValue, sToFind, compare)
  End If
  If (lInstr < mLength \ 2) Then
    Find = lInstr
  End If
End Function

'Public Function Replace(ByVal oldChar As Char, ByVal newChar As Char) As System.Text.StringBuilder
'Public Function Replace(ByVal oldChar As Char, ByVal newChar As Char, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
'Public Function Replace(ByVal oldValue As String, ByVal newValue As String) As System.Text.StringBuilder
'Public Function Replace(ByVal oldValue As String, ByVal newValue As String, ByVal startIndex As Integer, ByVal count As Integer) As System.Text.StringBuilder
Public Function Replace(ByVal oldValue As String, ByVal newValue As String, Optional ByVal startIndex As Long, Optional ByVal Count As Long) As StringBuilder
Dim PartStrVal As String, n As Long, StrPos As Long
Dim LStrVal As String, RStrVal As String
  n = Len(mStrValue)
  If startIndex > 0 And Count > 0 Then
    If startIndex < n Then
      If startIndex + Count > n Then Count = n - startIndex
      PartStrVal = MidB$(mStrValue, startIndex * 2, Count * 2)
    End If
  Else
    PartStrVal = mStrValue
  End If
  StrPos = InStr(1, PartStrVal, oldValue, vbBinaryCompare)
  If StrPos > 0 Then
    If Len(oldValue) >= Len(newValue) Then
      If Len(oldValue) = Len(newValue) Then
        
      Else
        
      End If
    Else
      LStrVal = LeftB$(mStrValue, StrPos)
    End If
    'Do While StrPos > 0
    '  StrPos = InStr(StrPos, PartStrVal, oldValue, vbBinaryCompare)
    'Loop
  End If
  Set Replace = Me
End Function


'Public Overrides Function ToString() As String
'Public Function ToString(ByVal startIndex As Integer, ByVal length As Integer) As String
Public Function ToString(Optional ByVal startIndex As Long, Optional ByVal length As Long) As String
Dim L As Long
'Set ToString = New cString
     ' The internal string:
  If LenB(mStrValue) > 0 Then
    If startIndex > 1 Then
      If length > 0 Then
        ToString = VBA.Mid$(mStrValue, startIndex, length)
      Else 'die Länge wurde nicht angegeben, also die gesamte verbleibende Länge
        ToString = VBA.Mid$(mStrValue, startIndex, length - startIndex)
      End If
    Else
      'Wenn es sich um einen Ansistring handelt, dann umwandeln
      ToString = StrConv(mStrValue, vbUnicode)
      If Mid$(ToString, 2, 1) = vbNullChar Then
        ToString = StrConv(ToString, vbFromUnicode)
      End If
      If mLength > 0 Then
        ToString = VBA.LeftB$(ToString, mLength)  '\2 durch 2, da mLength in Byte
      Else
        'das folgende kann sein, wenn mit SPtr für eine API-Funktion ein Zeiger
        'auf mStrValue übergeben wurde.
        'dann gibt es kein mLength, außerdem ist unbekannt, ob es sich um einen
        '2Byte- oder 1Byte per Char String handelt.
        L = InStr(1, ToString, vbNullChar, vbBinaryCompare) - 1
        If L > 0 Then
          ToString = VBA.Left$(ToString, L)
        Else
          L = InStr(1, mStrValue, vbNullChar, vbBinaryCompare) - 1
          If L > 0 Then
            ToString = VBA.Left$(mStrValue, L)
          End If
        End If
      End If
    End If
  End If
End Function
'Achtung Neue Funktion!!!!!!!! O.M. am
Public Function ToString0() As String
  ToString0 = VBA.LeftB$(mStrValue, (InStr(1, mStrValue, vbNullChar & vbNullChar, vbBinaryCompare) - 1) * 2)
End Function

Public Property Get SPtr() As Long
  SPtr = StrPtr(mStrValue)
End Property
Private Function NeedsAllocation(ByVal reqByteLength As Long) As Boolean
  NeedsAllocation = (LenB(mStrValue) <= reqByteLength)
End Function


'wird in .NET-Klasse nicht öffentlich verwendet
'Public Sub HeapMinimize()
'Dim iLen As Long
'
'   ' Reduce the string size so only the minimal chunks
'   ' are allocated:
'   If (m_iLen - m_iPos) > m_iChunkSize Then
'      iLen = m_iLen
'      Do While (iLen - m_iPos) > m_iChunkSize
'         iLen = iLen - m_iChunkSize
'      Loop
'      m_sString = Left$(m_sString, iLen \ 2)
'      m_iLen = iLen
'   End If
'
'End Sub
Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

'InternalEnsureCapacity (capacity);
'InternalEnsureCapacity(value);
'InternalEnsureCapacity (replace.Length + (_length - count))
'InternalEnsureCapacity (needed_cap);

'_str.InternalSetChar (index, value);
'_str.InternalSetChar (replaceIterate, newChar);
'_str.InternalSetChar(_length, value);

'_str.InternalSetLength(_length);

'String.InternalStrcpy (_str, startIndex, _str, startIndex + length, _length - (startIndex + length))
'String.InternalStrcpy (_str, startIndex, replace);
'String.InternalStrcpy (_str, startIndex + replace.Length, end);
'String.InternalStrcpy (_str, _length, value);
'String.InternalStrcpy (_str, _length, value);
'String.InternalStrcpy (_str, _length, value, startIndex, charCount);
'internal extern static void
'InternalStrcpy (String dest, int destPos, String src);
'InternalStrcpy (String dest, int destPos, char[] chars);
'InternalStrcpy (String dest, int destPos, String src, int sPos, int count);
'InternalStrcpy (String dest, int destPos, char[] chars, int sPos, int count);
'
Private Sub InternalStrcpy(dest As String, ByVal destPos As Long, src As String, Optional ByVal sPos As Long, Optional ByVal Count As Long)
'kopiert den String src an die Stelle destPos im string dest
'kopiert den Teilstring src = mid$(src, spos, count) an die Stelle destPos im String dest
Dim lLen As Long
Dim PartStr As String
  If Count = 0 Then Count = LenB(src)
  If sPos = 0 Then sPos = 1
  PartStr = Mid$(src, sPos, Count)
  'mLength = Len(PartStr)
  lLen = 2 * Len(PartStr) 'mLength
  'wie errechnet sich jetzt die Länge des Strings?
  '
  'mLength = mLength + LenB(PartStr)
' CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
  
  CopyMemory ByVal UnsignedAdd(StrPtr(dest), destPos), ByVal StrPtr(PartStr), lLen
  'MidB$(dest, destPos) = MidB$(src, sPos, Count)
End Sub

'string tmp = String.InternalAllocateStr (capacity);
Private Function InternalAllocateStr(ByVal mCap As Long) As String
  'sollte man besser Space$ oder String$(Length, vbNullString) verwenden??
  '->siehe AllocString auf www.xbeat.net/vbspeed/c_AllocString.htm
  'InternalAllocateStr = Space$(mCap)
  Call CopyMemory(ByVal VarPtr(InternalAllocateStr), SysAllocStringByteLen(0&, mCap), 4&) ' + mCap), 4&)
  'InternalAllocateStr = SysAllocStringLen(ByVal 0&, mCap \ 2)
End Function

Private Sub InternalEnsureCapacity(ByVal size As Long)
Dim tmp As String
  If (size > mLength Or mCachedStr = mStrValue) Then
    mCapacity = mLength

    '// Try double buffer, if that doesn't work, set the length as capacity
    If (size > mCapacity) Then
          
      '// The first time a string is appended, we just set _cached_str
      '// and _str to it. This allows us to do some optimizations.
      '// Below, we take this into account.
      If (mCachedStr = mStrValue) And (mCapacity < constDefaultCapacity) Then
        mCapacity = constDefaultCapacity
      End If
      
      'capacity = capacity << 1; ShiftLeft????
      If (size > mCapacity) Then mCapacity = size
      If (mCapacity >= 2147483647 Or mCapacity < 0) Then
        mCapacity = 2147483647 'Int32.MaxValue;
      End If

      If (mCapacity > mMaxCapacity And size <= mMaxCapacity) Then
        mCapacity = mMaxCapacity
      End If
      
      If (mCapacity > mMaxCapacity) Then
      '  throw new ArgumentOutOfRangeException ("size", "capacity was less than the current size.");
      End If
    End If

    tmp = InternalAllocateStr(mCapacity)
    If (mLength > 0) Then
      Call InternalStrcpy(tmp, 0, mStrValue, 0, mLength)
    End If
    mStrValue = tmp
  End If

  mCachedStr = vbNullString
End Sub

Public Property Let length(LngVal As Long)
  mLength = LngVal * 2
  If mLength = 0 Then
    Call NewC
  End If
End Property
Public Property Get length() As Long
  length = mLength \ 2
End Property

Private Sub Class_Terminate()
  'Length = 0
  'mStrValue = vbNullString
  mStrValue = InternalAllocateStr(constDefaultCapacity)
End Sub

