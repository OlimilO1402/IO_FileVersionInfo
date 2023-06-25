VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FileVersionInfo"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtFileName 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton BtnFileVersion 
      Caption         =   "GetVersionInfo"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "?"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtFileVersionInfo 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FVInfo As FileVersionInfo

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
           App.FileDescription & vbCrLf & "Hint: " & App.Comments, vbInformation
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    TxtFileName.Text = "C:\Windows\System32\kernel32.dll"
    'TxtFileName.Text = "C:\Windows\System32\msvcr100.dll"
End Sub

Private Sub BtnFileVersion_Click()
Try: On Error GoTo Catch
    Set FVInfo = MNew.FileVersionInfo(TxtFileName.Text)
    If FVInfo Is Nothing Then Exit Sub
    TxtFileVersionInfo.Text = FVInfo.ToStr
    Exit Sub
Catch:
    MsgBox ("Probably file not found")
End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    Dim Brdr As Single ': Brdr = 8 * 15
    l = TxtFileName.Left
    T = TxtFileName.Top
    W = Form1.ScaleWidth - Brdr - l
    H = TxtFileName.Height
    If ((W > 0) And (H > 0)) Then Call TxtFileName.Move(l, T, W, H)
    l = TxtFileVersionInfo.Left
    T = TxtFileVersionInfo.Top
    W = Form1.ScaleWidth - Brdr - l
    H = Form1.ScaleHeight - Brdr - T
    If ((W > 0) And (H > 0)) Then Call TxtFileVersionInfo.Move(l, T, W, H)
End Sub

Private Sub TxtFileName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then
        BtnFileVersion_Click
    End If
End Sub

Private Sub TxtFileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TxtOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub TxtFileVersionInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TxtOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub TxtOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count = 0 Then Exit Sub
    Dim FNm As String: FNm = Data.Files(1)
    TxtFileName.Text = FNm
    BtnFileVersion_Click
End Sub
