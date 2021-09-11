VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VS_FIXEDFILEINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Achtung in LutzRoeders .net-Reflector werden die Memberfelder
'in alphabetischer Reihenfolge ausgegeben. Dies ist jedoch die
'falsche Anordnung zur Verwendung für API-Funktionen (meistens)
'im Reflector angegeben ist:
' Fields
'Public dwFileDateLS As Integer
'Public dwFileDateMS As Integer
'Public dwFileFlags As Integer
'Public dwFileFlagsMask As Integer
'Public dwFileOS As Integer
'Public dwFileSubtype As Integer
'Public dwFileType As Integer
'Public dwFileVersionLS As Integer
'Public dwFileVersionMS As Integer
'Public dwProductVersionLS As Integer
'Public dwProductVersionMS As Integer
'Public dwSignature As Integer
'Public dwStructVersion As Integer
'
'die richtige Reihenfolge ist wichtig:
Public dwSignature As Long          '
Public dwStructVersion As Long      '
Public dwFileVersionMS As Long      '
Public dwFileVersionLS As Long      '
Public dwProductVersionMS As Long   '
Public dwProductVersionLS As Long   '
Public dwFileFlagsMask As Long      '
Public dwFileFlags As Long          '
Public dwFileOS As Long             '
Public dwFileType As Long           '
Public dwFileSubtype As Long        '
Public dwFileDateMS As Long         '
Public dwFileDateLS As Long         '

Public Property Get value() As Long
  value = VarPtr(dwSignature)
End Property
Public Property Get ByteLength() As Long
  ByteLength = 52 ' = 13 * 4
End Property

