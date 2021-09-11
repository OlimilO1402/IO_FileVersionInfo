Imports System.Diagnostics
Imports System.Runtime.InteropServices
Public Class Form1
  Inherits System.Windows.Forms.Form

#Region " Vom Windows Form Designer generierter Code "

  Public Sub New()
    MyBase.New()

    ' Dieser Aufruf ist für den Windows Form-Designer erforderlich.
    InitializeComponent()

    ' Initialisierungen nach dem Aufruf InitializeComponent() hinzufügen

  End Sub

  ' Die Form überschreibt den Löschvorgang der Basisklasse, um Komponenten zu bereinigen.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  ' Für Windows Form-Designer erforderlich
  Private components As System.ComponentModel.IContainer

  'HINWEIS: Die folgende Prozedur ist für den Windows Form-Designer erforderlich
  'Sie kann mit dem Windows Form-Designer modifiziert werden.
  'Verwenden Sie nicht den Code-Editor zur Bearbeitung.
  Friend WithEvents BtnFileVersion As System.Windows.Forms.Button
  Friend WithEvents TxtFileVersionInfo As System.Windows.Forms.TextBox
  Friend WithEvents TxtFileName As System.Windows.Forms.TextBox
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.TxtFileVersionInfo = New System.Windows.Forms.TextBox
    Me.TxtFileName = New System.Windows.Forms.TextBox
    Me.BtnFileVersion = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'TxtFileVersionInfo
    '
    Me.TxtFileVersionInfo.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.TxtFileVersionInfo.Location = New System.Drawing.Point(120, 88)
    Me.TxtFileVersionInfo.Multiline = True
    Me.TxtFileVersionInfo.Name = "TxtFileVersionInfo"
    Me.TxtFileVersionInfo.Size = New System.Drawing.Size(312, 184)
    Me.TxtFileVersionInfo.TabIndex = 0
    Me.TxtFileVersionInfo.Text = ""
    '
    'TxtFileName
    '
    Me.TxtFileName.Location = New System.Drawing.Point(120, 48)
    Me.TxtFileName.Name = "TxtFileName"
    Me.TxtFileName.Size = New System.Drawing.Size(288, 20)
    Me.TxtFileName.TabIndex = 1
    Me.TxtFileName.Text = ""
    '
    'BtnFileVersion
    '
    Me.BtnFileVersion.Location = New System.Drawing.Point(8, 8)
    Me.BtnFileVersion.Name = "BtnFileVersion"
    Me.BtnFileVersion.Size = New System.Drawing.Size(104, 24)
    Me.BtnFileVersion.TabIndex = 2
    Me.BtnFileVersion.Text = "GetVersionInfo"
    '
    'Form1
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(528, 286)
    Me.Controls.Add(Me.BtnFileVersion)
    Me.Controls.Add(Me.TxtFileName)
    Me.Controls.Add(Me.TxtFileVersionInfo)
    Me.Name = "Form1"
    Me.Text = "Test FileVersionInfo"
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    TxtFileName.Text = "C:\Windows\System32\kernel32.dll"
  End Sub

  Private Sub BtnFileVersion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFileVersion.Click
    Dim VI As FileVersionInfo
    Try
      VI = FileVersionInfo.GetVersionInfo(TxtFileName.Text)
      TxtFileVersionInfo.Text = VI.ToString
      MsgBox(TxtFileVersionInfo.Text.Length)
    Catch
      MsgBox("Probably File Not Found")
    End Try
  End Sub

  Private Sub TxtFileName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFileName.KeyUp
    If e.KeyCode = Keys.Enter Then
      BtnFileVersion_Click(sender, New System.EventArgs)
    End If
  End Sub

End Class
