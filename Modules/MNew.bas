Attribute VB_Name = "MNew"
Option Explicit

Public Function FileVersionInfo(aPFN As String) As FileVersionInfo
    Set FileVersionInfo = New FileVersionInfo: FileVersionInfo.New_ aPFN
End Function
