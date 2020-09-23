Attribute VB_Name = "CleanIT"
'********************************************************************
' The cleaning tools                                                *
'********************************************************************

Option Explicit

Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long

' Clears Typed URLs
Public Function CleanIE()
    Dim URL_List As String
    Dim RegPath As String
    Dim Icount As Integer
    
    On Error Resume Next
    RegPath = "Software\Microsoft\Internet Explorer\TypedURLs"
    ' You must close Internet Explorer first
    Icount = 1
    URL_List = getstring(HKEY_CURRENT_USER, RegPath, "url" & Icount)
    Do Until Len(URL_List) = 0
        DeleteValue HKEY_CURRENT_USER, RegPath, "url" & Icount
        Icount = Icount + 1
        URL_List = getstring(HKEY_CURRENT_USER, RegPath, "url" & Icount)
    Loop
End Function

' Clear Run Dialog
Public Function CleanRunMnu()
    Dim Mnu_List As String
    Dim Menu_Item As String
    Dim Icount As Integer
    Dim RegPath As String
    
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    On Error Resume Next
    Mnu_List = getstring(HKEY_CURRENT_USER, RegPath, "MRUList")
    For Icount = 1 To Len(Mnu_List)
        Menu_Item = Mid(Mnu_List, Icount, 1)
        DeleteValue HKEY_CURRENT_USER, RegPath, Menu_Item
    Next
    DeleteValue HKEY_CURRENT_USER, RegPath, "MRUList"
End Function

' Clean search dialog
Public Function CleanFind()
    Dim Mnu_List As String
    Dim Menu_Item As String
    Dim Icount As Integer
    Dim RegPath As String
    
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Doc Find Spec MRU"
    On Error Resume Next
    Mnu_List = getstring(HKEY_CURRENT_USER, RegPath, "MRUList")
    For Icount = 1 To Len(Mnu_List)
        Menu_Item = Mid(Mnu_List, Icount, 1)
        DeleteValue HKEY_CURRENT_USER, RegPath, Menu_Item
    Next
    DeleteValue HKEY_CURRENT_USER, RegPath, "MRUList"
End Function

' Clean search for computers
Public Function CleanFindComp()
    Dim Mnu_List As String
    Dim Menu_Item As String
    Dim Icount As Integer
    Dim RegPath As String
    
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\FindComputerMRU"
    On Error Resume Next
    Mnu_List = getstring(HKEY_CURRENT_USER, RegPath, "MRUList")
    For Icount = 1 To Len(Mnu_List)
        Menu_Item = Mid(Mnu_List, Icount, 1)
        DeleteValue HKEY_CURRENT_USER, RegPath, Menu_Item
    Next
    DeleteValue HKEY_CURRENT_USER, RegPath, "MRUList"
End Function

' Clean Recent Documents
Public Function CleanDocs()
    SHAddToRecentDocs 0, 0
    CleanDX
End Function

' Recent Documents dword MRU
Private Function CleanDX()
    Dim Mnu_List As String
    Dim Menu_Item As String
    Dim Icount As Integer
    Dim RegPath As String
    
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
    On Error Resume Next
    Mnu_List = getstring(HKEY_CURRENT_USER, RegPath, "MRUList")
    For Icount = 1 To Len(Mnu_List)
        Menu_Item = Mid(Mnu_List, Icount, 1)
        DeleteValue HKEY_CURRENT_USER, RegPath, Menu_Item
    Next
    DeleteValue HKEY_CURRENT_USER, RegPath, "MRUList"
End Function

' Clean Windows Media Player
Public Function CleanMediaPlayer()
    Dim RegPath As String
    Dim Icount As Integer
    Dim File_List As String
    
    On Error Resume Next
    RegPath = "Software\Microsoft\MediaPlayer\Player\RecentFileList"
    Icount = 1
    File_List = getstring(HKEY_CURRENT_USER, RegPath, "File" & Icount)
    Do Until Len(File_List) = 0
        DeleteValue HKEY_CURRENT_USER, RegPath, "File" & Icount
        Icount = Icount + 1
        File_List = getstring(HKEY_CURRENT_USER, RegPath, "File" & Icount)
    Loop
End Function

' Clean Windows Paint Brush
Public Function CleanPaintBrushFileLst()
    Dim RegPath As String
    Dim Icount As Integer
    Dim File_List As String
    
    On Error Resume Next
    RegPath = "Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List"
    Icount = 1
    File_List = getstring(HKEY_CURRENT_USER, RegPath, "File" & Icount)
    Do Until Len(File_List) = 0
        DeleteValue HKEY_CURRENT_USER, RegPath, "File" & Icount
        Icount = Icount + 1
        File_List = getstring(HKEY_CURRENT_USER, RegPath, "File" & Icount)
    Loop
End Function
