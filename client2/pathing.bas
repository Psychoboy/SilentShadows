Attribute VB_Name = "pathing"
Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Function SystemDir() As String
    Dim Gwdvar As String, Gwdvar_Length As Integer
    Gwdvar = Space(255)
    Gwdvar_Length = GetSystemDirectory(Gwdvar, 255)
    SystemDir = Left(Gwdvar, Gwdvar_Length)
End Function
Function WindowsDir() As String
    Dim Gwdvar As String, Gwdvar_Length As Integer
    Gwdvar = Space(255)
    Gwdvar_Length = GetWindowsDirectory(Gwdvar, 255)
    WindowsDir = Left(Gwdvar, Gwdvar_Length)
End Function


