VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnFileVersion 
      Caption         =   "GetVersionInfo"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtFileVersionInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
   Begin VB.TextBox TxtFileName 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FileVersionInfo As New FileVersionInfo

Private Sub Form_Load()
  TxtFileName.Text = "C:\Windows\System32\kernel32.dll"
End Sub

Private Sub BtnFileVersion_Click()
Dim VI As FileVersionInfo
TryE: On Error GoTo CatchE
  Set VI = FileVersionInfo.GetVersionInfo(TxtFileName.Text)
  TxtFileVersionInfo.Text = VI.ToString
  Exit Sub
CatchE:
  MsgBox ("Probably file not found")
End_Try:
End Sub

Private Sub Form_Resize()
Dim L As Single, T As Single, W As Single, H As Single
Dim Brdr As Single: Brdr = 8 * 15
  L = TxtFileName.Left
  T = TxtFileName.Top
  W = Form1.ScaleWidth - Brdr - L
  H = TxtFileName.Height
  If ((W > 0) And (H > 0)) Then Call TxtFileName.Move(L, T, W, H)
  L = TxtFileVersionInfo.Left
  T = TxtFileVersionInfo.Top
  W = Form1.ScaleWidth - Brdr - L
  H = Form1.ScaleHeight - Brdr - T
  If ((W > 0) And (H > 0)) Then Call TxtFileVersionInfo.Move(L, T, W, H)
End Sub

Private Sub TxtFileName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbEnter Then
    BtnFileVersion_Click
  End If
End Sub
