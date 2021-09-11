VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnFileVersion 
      Caption         =   "GetVersionInfo"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3015
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
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox TxtFileName 
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3015
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

    BtnFileVersion.Caption = "GetVersionInfo"
    Call BtnFileVersion.Move(8, 8, 105, 25)
    Form1.Width = 521 * 15: Form1.Height = 403 * 15
    TxtFileName.Text = "C:\Windows\System32\kernel32.dll"
    'TxtFileVersionInfo.MultiLine = True

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

Private Sub TxtFileName_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbEnter Then
        BtnFileVersion_Click
    End If

End Sub

Private Sub Form_Resize()

Dim L As Single, T As Single, W As Single, H As Single
    
    L = BtnFileVersion.Left + BtnFileVersion.Width
    T = BtnFileVersion.Top
    W = Form1.ScaleWidth - 8 - L
    H = 19
    If ((W > 0) And (H > 0)) Then _
        Call TxtFileName.Move(L, T, W, H)
    L = 8
    T = BtnFileVersion.Top + BtnFileVersion.Height + 8
    W = Form1.ScaleWidth - 8 - L
    H = Form1.ScaleHeight - 8 - T
    If ((W > 0) And (H > 0)) Then _
        Call TxtFileVersionInfo.Move(L, T, W, H)

End Sub

