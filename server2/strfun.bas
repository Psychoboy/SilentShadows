Attribute VB_Name = "Module1"
'Option Compare Database
Option Explicit

'String Functions Library
'This library contains various VBA functions for often-used string handling needs
'Usage note: All string comparisons use binary comparison, i.e. "A" <> "a"
'Copyright 2000 Roman Koch (roman@romankoch.ch)

Public Const Blank As String = " "

Function SF_count(ByVal Haystack As String, ByVal Needle As String) As Long


    'count the number of occurences of needle in haystack
    'SF_count("   This is   my string   ","i") returns 3
    Dim i As Long, j As Long
    If SF_isNothing(Needle) Then
        SF_count = 0
    Else
        i = InStr(1, Haystack, Needle, vbBinaryCompare)
        If i = 0 Then
            SF_count = 0
        Else
            i = 0
            For j = 1 To Len(Haystack)
                If Mid(Haystack, j, Len(Needle)) = Needle Then i = i + 1
            Next j
            SF_count = i
        End If
    End If
End Function

Function SF_countWords(ByVal Haystack As String) As Long


    'count number of words in a string
    'if the string is empty, 0 is returned
    'SF_countWords("   This is   my string   ") returns 4
    Dim strChar As String
    Dim lngCount As Long, i As Long
    Haystack = SF_unSpace(Haystack)
    If SF_isNothing(Haystack) Then
        SF_countWords = 0
    Else
        lngCount = 1
        For i = 1 To Len(Haystack)
            strChar = Mid(Haystack, i, 1)
            If strChar = Blank Then
                lngCount = lngCount + 1
            End If
        Next i
        SF_countWords = lngCount
    End If
End Function

Function SF_getWord(ByVal Haystack As String, ByVal WordNumber As Long) As String


    'return nth word of a string
    'if the string is empty, a zero-length string is returned
    'if there is only one word, the initial string is returned
    'if wordnumber is 0 or negative, a zero-length string is returned
    'if wordnumber is larger than the number of words in the string, a zero-length string is returned
    'SF_getWord("   This is   my string   ",2) returns "is"
    Dim i, lngWords As Long
    Haystack = SF_unSpace(Haystack)
    If SF_isNothing(Haystack) Then
        SF_getWord = Haystack
    Else
        If WordNumber > 0 Then
            lngWords = SF_countWords(Haystack)
            If WordNumber > lngWords Then
                Haystack = ""
            Else
                If lngWords > 1 Then
                    'cut words at the left
                    For i = 1 To WordNumber - 1
                        Haystack = Mid(Haystack, InStr(Haystack, Blank) + 1)
                    Next i
                    'cut words at the right, if any
                    i = InStr(Haystack, Blank)
                    If i > 0 Then Haystack = Left(Haystack, i - 1)
                End If
            End If
        Else
            Haystack = ""
        End If
        SF_getWord = Haystack
    End If
End Function

Function SF_InstrRev(ByVal Haystack As String, ByVal Needle As String) As Long


    'find the last occurence of needle in haystack (the VB instr function finds the first occurence)
    'SF_InstrRev("   This is   my string   ","i") = 20
    Dim i As Long, j As Long
    i = InStr(1, Haystack, Needle, vbBinaryCompare)
    If SF_isNothing(Needle) Then
        SF_InstrRev = 0
    Else
        If i = 0 Then
            SF_InstrRev = 0
        Else
            If StrComp(Needle, Haystack, vbBinaryCompare) = 0 Then
                SF_InstrRev = 1
            Else
                For j = Len(Haystack) To 1 Step -1
                    i = InStr(j, Haystack, Needle, vbBinaryCompare)
                    If i > 0 Then
                        SF_InstrRev = i
                        Exit Function
                    End If
                Next j
            End If
        End If
    End If
End Function

Function SF_isNothing(ByVal Haystack As String) As Boolean


    'check if there is anything in a string (to avoid testing for
    'isnull, isempty, and zero-length strings)
    'SF_isNothing("   This is   my string   ") returns False
    If Haystack & "" = "" Then
        SF_isNothing = True
    Else
        SF_isNothing = False
    End If
End Function

Function SF_remove(ByVal Haystack As String, ByVal Needle As String) As String


    'remove first occurence of needle in haystack
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, a zero-length string is returned
    'SF_remove("   This is   my string   ","   This is   m") returns "y string   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_remove = Haystack
    Else
        i = InStr(1, Haystack, Needle, vbBinaryCompare)
        If i = 0 Then
            SF_remove = Haystack
        Else
            SF_remove = SF_splitLeft(Haystack, Needle) & SF_splitRight(Haystack, Needle)
        End If
    End If
End Function

Function SF_removeRev(ByVal Haystack As String, ByVal Needle As String) As String


    'remove last occurence of needle in haystack
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, a zero-length string is returned
    'SF_removeRev("   This is   my string   ","i") returns "   This is   my strng   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_removeRev = Haystack
    Else
        i = SF_InstrRev(Haystack, Needle)
        If i = 0 Then
            SF_removeRev = Haystack
        Else
            SF_removeRev = Left(Haystack, i - 1) & Mid(Haystack, i + Len(Needle))
        End If
    End If

End Function

Function SF_removeAllOnce(ByVal Haystack As String, ByVal Needle As String) As String


    'remove all occurrences of needle in haystack exactly once
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, a zero-length string is returned
    'SF_removeAllOnce("1122a1122","12") returns "12a12"
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_removeAllOnce = Haystack
    Else
        i = InStr(1, Haystack, Needle, vbBinaryCompare)
        Do While i > 0
            Haystack = Left(Haystack, i - 1) & Mid(Haystack, i + Len(Needle))
            i = InStr(i, Haystack, Needle, vbBinaryCompare)
        Loop
        SF_removeAllOnce = Haystack
    End If
End Function

Function SF_removeAll(ByVal Haystack As String, ByVal Needle As String) As String


    'remove all occurrences of needle in haystack, even those created during removal
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, a zero-length string is returned
    'SF_removeAll("1122a1122","12") returns "a"
    Do While InStr(1, Haystack, Needle) > 0
        Haystack = SF_removeAllOnce(Haystack, Needle)
    Loop
    SF_removeAll = Haystack
End Function

Function SF_replace(ByVal Haystack As String, ByVal Needle As String, ByVal NewNeedle As String) As String


    'replace first occurence of needle in haystack with newneedle
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, newneedle is returned
    'if needle is equal to newneedle, haystack is returned
    'SF_replace("   This is   my string   ","my","your") returns "   This is   your string   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_replace = Haystack
    Else
        If StrComp(Needle, NewNeedle, vbBinaryCompare) = 0 Then
            SF_replace = Haystack
        Else
            i = InStr(1, Haystack, Needle, vbBinaryCompare)
            If i = 0 Then
                SF_replace = Haystack
            Else
                SF_replace = SF_splitLeft(Haystack, Needle) & NewNeedle & SF_splitRight(Haystack, Needle)
            End If
        End If
    End If
End Function

Function sf_replaceRev(ByVal Haystack As String, ByVal Needle As String, ByVal NewNeedle As String) As String


    'replace last occurence of needle in haystack with newneedle
    'if needle is empty or not found, haystack is returned
    'if needle is equal to haystack, newneedle is returned
    'if needle is equal to newneedle, haystack is returned
    'SF_replaceRev("   This is   my string   ","i","o") returns "   This is   my strong   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        sf_replaceRev = Haystack
    Else
        If StrComp(Needle, NewNeedle, vbBinaryCompare) = 0 Then
            sf_replaceRev = Haystack
        Else
            i = SF_InstrRev(Haystack, Needle)
            If i = 0 Then
                sf_replaceRev = Haystack
            Else
                sf_replaceRev = Left(Haystack, i - 1) & NewNeedle & Mid(Haystack, i + Len(Needle))
            End If
        End If
    End If
End Function

Function SF_replaceAllOnce(ByVal Haystack As String, ByVal Needle As String, ByVal NewNeedle As String) As String


    'replace all occurrences of needle in haystack with newneedle exactly once
    'if needle is empty or not found, haystack is returned
    'if needle is equal to newneedle, haystack is returned
    'if needle is equal to haystack, newneedle is returned
    'SF_replaceAllOnce("   This is   my string   ","i","ee") returns "   Thees ees   my streeng   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_replaceAllOnce = Haystack
    Else
        If StrComp(Needle, NewNeedle, vbBinaryCompare) = 0 Then
            SF_replaceAllOnce = Haystack
        Else
            i = InStr(1, Haystack, Needle, vbBinaryCompare)
            Do While i > 0
                Haystack = Left(Haystack, i - 1) & NewNeedle & Mid(Haystack, i + Len(Needle))
                i = i + Len(NewNeedle)
                i = InStr(i, Haystack, Needle, vbBinaryCompare)
            Loop
        SF_replaceAllOnce = Haystack
        End If
    End If
End Function

Function SF_replaceAll(ByVal Haystack As String, ByVal Needle As String, ByVal NewNeedle As String) As String


    'replace all occurrences of needle in haystack with newneedle, even those created during replacing
    'if needle is empty or not found, haystack is returned
    'if needle is equal to newneedle, haystack is returned
    'if needle is equal to haystack, newneedle is returned
    'if needle is a subset of newneedle, the function would loop;
    'to avoid this, SF_replaceAllOnce is executed instead
    'SF_replaceAll("   This is   my string   ","i","ee") returns "   Thees ees   my streeng   "
    If InStr(1, NewNeedle, Needle, vbBinaryCompare) > 0 Then
        Haystack = SF_replaceAllOnce(Haystack, Needle, NewNeedle)
    Else
        Do While InStr(1, Haystack, Needle, vbBinaryCompare) > 0
            Haystack = SF_replaceAllOnce(Haystack, Needle, NewNeedle)
        Loop
    End If
    SF_replaceAll = Haystack
End Function

Function SF_splitLeft(ByVal Haystack As String, ByVal Needle As String) As String


    'return left part of haystack delimited by the first occurrence of needle
    'if needle is empty or not found, haystack is returned
    'if haystack starts with needle (or is equal to needle), a zero-length string is returned
    'SF_splitLeft("   This is   my string   ","s is") returns "   Thi"
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_splitLeft = Haystack
    Else
        i = InStr(1, Haystack, Needle, vbBinaryCompare)
        If i = 0 Then
            SF_splitLeft = Haystack
        Else
            SF_splitLeft = Left(Haystack, i - 1)
        End If
    End If
End Function

Function SF_splitRight(ByVal Haystack As String, ByVal Needle As String) As String


    'return right part of haystack delimited by the first occurrence of needle
    'if needle is empty or not found, haystack is returned
    'if haystack ends with needle (or is equal to needle), a zero-length string is returned
    'SF_splitRight("   This is   my string   "," my s") returns "tring   "
    Dim i As Long
    If SF_isNothing(Needle) Then
        SF_splitRight = Haystack
    Else
        i = InStr(1, Haystack, Needle, vbBinaryCompare)
        If i = 0 Then
            SF_splitRight = Haystack
        Else
            SF_splitRight = Mid(Haystack, i + Len(Needle))
        End If
    End If
End Function

Function SF_unSpace(ByVal Haystack As String) As String


    'remove duplicate blanks in a string
    'SF_unspace("   This is   my string   ") returns "This is my string"
    If SF_isNothing(Haystack) Then
        SF_unSpace = Haystack
    Else
        Haystack = Trim(Haystack)
        Do While InStr(Haystack, Blank & Blank) > 0
            Haystack = SF_replaceAllOnce(Haystack, Blank & Blank, Blank)
        Loop
        SF_unSpace = Haystack
    End If
End Function

Sub SF_Test()

    
    'test the string functions
    Dim Haystack As String, Needle As String, NewNeedle As String, strWork As String
    Dim NumericParm As Long
    Haystack = InputBox("Enter the Haystack string", "SF_Test Input 1 of 4")
    Needle = InputBox("Enter the Needle string", "SF_Test Input 2 of 4")
    NewNeedle = InputBox("Enter the NewNeedle string", "SF_Test Input 3 of 4")
    strWork = InputBox("Enter a numeric parameter", "SF_Test Input 4 of 4")
    If IsNumeric(strWork) Then
        NumericParm = CLng(strWork)
    Else
        NumericParm = 0
    End If
    strWork = "Parameters:" & vbCr & _
            "Haystack: [" & Haystack & "]" & vbCr & _
            "Needle: [" & Needle & "]" & vbCr & _
            "NewNeedle: [" & NewNeedle & "]" & vbCr & _
            "Number: [" & NumericParm & "]" & vbCr & vbCr & _
            "Results:" & vbCr & _
            "1. SF_count: [" & SF_count(Haystack, Needle) & "]" & vbCr & _
            "2. SF_countWords: [" & SF_countWords(Haystack) & "]" & vbCr & _
            "3. SF_getWord: [" & SF_getWord(Haystack, NumericParm) & "]" & vbCr & _
            "4. SF_InstrRev: [" & SF_InstrRev(Haystack, Needle) & "]" & vbCr & _
            "5. SF_isNothing: [" & SF_isNothing(Haystack) & "]" & vbCr & _
            "6. SF_remove: [" & SF_remove(Haystack, Needle) & "]" & vbCr & _
            "7. SF_removeRev: [" & SF_removeRev(Haystack, Needle) & "]" & vbCr & _
            "8. SF_removeAllOnce: [" & SF_removeAllOnce(Haystack, Needle) & "]" & vbCr & _
            "9. SF_removeAll: [" & SF_removeAll(Haystack, Needle) & "]" & vbCr & _
            "10. SF_replace: [" & SF_replace(Haystack, Needle, NewNeedle) & "]" & vbCr & _
            "11. SF_replaceRev: [" & sf_replaceRev(Haystack, Needle, NewNeedle) & "]" & vbCr & _
            "12. SF_replaceAllOnce: [" & SF_replaceAllOnce(Haystack, Needle, NewNeedle) & "]" & vbCr & _
            "13. SF_replaceAll: [" & SF_replaceAll(Haystack, Needle, NewNeedle) & "]" & vbCr & _
            "14. SF_splitLeft: [" & SF_splitLeft(Haystack, Needle) & "]" & vbCr & _
            "15. SF_splitRight: [" & SF_splitRight(Haystack, Needle) & "]" & vbCr & _
            "16. SF_unSpace: [" & SF_unSpace(Haystack) & "]"
    MsgBox strWork, vbOKOnly, "Results"
End Sub
