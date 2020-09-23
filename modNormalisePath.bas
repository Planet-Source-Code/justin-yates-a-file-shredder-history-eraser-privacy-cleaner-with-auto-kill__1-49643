Attribute VB_Name = "modNormalisePath"
Option Explicit

' If path does not have a "\" at the end of the path, then add one
Public Function NormalisePath(ByVal strPath As String) As String
    If Right$(strPath, 1) = "\" Then
        NormalisePath = strPath
    Else
        NormalisePath = strPath & "\"
    End If
End Function
