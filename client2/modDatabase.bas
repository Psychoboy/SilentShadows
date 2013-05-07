Attribute VB_Name = "modDatabase"
Option Explicit

Function FileExist(ByVal filename As String) As Boolean
    If Dir(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Function FileExist2(ByVal filename As String) As Boolean
    If Dir(filename) = "" Then
        FileExist2 = False
    Else
        FileExist2 = True
    End If
End Function

Sub AddLog(ByVal Text As String)
Dim filename As String
Dim f As Long

    If Trim(Command) = "-debug" Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If
        
        filename = App.Path & "\debug.txt"
    
        If Not FileExist("debug.txt") Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open filename For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub SaveLocalMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
            
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , SaveMap
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , SaveMap
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
Dim filename As String
Dim f As Long
Dim TmpMap As MapRec

    filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString(Header, Var, value, File)
End Sub

Public Sub updateBList()
frmMirage.BList.Clear

Dim i As Long

    For i = 1 To 10
    frmMirage.BList.AddItem (BuDDy(i))
    
Next i
End Sub
