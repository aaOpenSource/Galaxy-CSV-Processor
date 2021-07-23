#If Win64 Then
  Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                      (ByVal IpBuffer As String, nSize As Long) As Long
  Public Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                      (ByVal lpBuffer As String, nSize As Long) As Long

#Else
  Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA"
                      (ByVal IpBuffer As String, nSize As Long) As Long
  Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                      (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Sub ShowForm()
    frmMain.Show
End Sub
