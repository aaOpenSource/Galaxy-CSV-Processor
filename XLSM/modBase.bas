Attribute VB_Name = "modBase"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub ShowForm()
    frmMain.Show
End Sub
