Attribute VB_Name = "CheckProcess"
' By: Anthony Christianson
'
' Inputs:Send an EXE Name, ie explorer.e
'     xe
'
' Returns:it populates a Running variabl
'     e that is updated when the exe name prop
'     erty is set allowing you to capture if t
'     hat exe is running or not
'
' Assumes:DOES NOT WORK ON Windows 95 or
'     98 or ME. Assume you know how to use a V
'     B Class. I used this class in a game ser
'     ver application that would detect when m
'     y server crashed then restart it.
'
' Side Effects:No Known Side Effects
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=37202&lngWId=1'for details.'**************************************

Option Explicit
'Majority of this code was taken from
'http://www.allapi.net/apilist/example.p
'     hp?example=Enumerate%20Processes
'AllAPI.net
'A great place to go for code
'Version 2
'Added the Token priviledges to see all
'     processes whether under your context
'or another


Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long


Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long


Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long


Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long


Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long


Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long


Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long


Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
    Private Const PROCESS_QUERY_INFORMATION = 1024
    Private Const PROCESS_VM_READ = 16
    Private Const MAX_PATH = 260
    Private Const TOKEN_ADJUST_PRIVILEGES = &H20
    Private Const TOKEN_QUERY = &H8
    Private Const SE_PRIVILEGE_ENABLED = &H2
    Private mvar_Exename As String
    Private mvar_FullPath As String
    Private mvar_WorkDir As String


Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
    End Type


Private Type LUID_AND_ATTRIBUTES
    TheLuid As LUID
    Attributes As Long
    End Type


Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Public Sub StartProcess(ByVal cmdLine As String)
    Print #5, "Changing Drive"
    ChDrive Left(mvar_FullPath, 1)
    Print #5, "Changing Directory"
    ChDir mvar_WorkDir


    If Len(cmdLine) > 0 Then
        Print #5, "Shell"
        Shell "cmd /c " & mvar_FullPath & " " & cmdLine, vbHide
    Else
        Shell "cmd /c " & mvar_FullPath & " " & cmdLine, vbHide
    End If
End Sub

Public Function GetProcessStatus(ByVal EXEName As String) As Boolean
    Dim lngLength As Long, lngCBSize As Long
    Dim lngCBSizeReturned As Long, strProcessName As String
    Dim lngNumElements As Long, lngCBSize2 As Long
    Dim lngLoop As Long, lngReturn As Long
    Dim lngSize As Long, lngHwndProcess As Long
    Dim lngProcessIDs() As Long, lngModules(1 To 200) As Long
    Dim strModuleName As String
    Dim arrSplit As Variant
    GetProcessStatus = False
    EXEName = UCase$(Trim$(EXEName))
    lngLength = Len(EXEName)
    lngCBSize = 8
    lngCBSizeReturned = 96


    Do While lngCBSize <= lngCBSizeReturned


        DoEvents
            lngCBSize = lngCBSize * 2
            ReDim lngProcessIDs(lngCBSize / 4) As Long
            lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
        Loop
        lngNumElements = lngCBSizeReturned / 4


        For lngLoop = 1 To lngNumElements


            DoEvents
                lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))


                If lngHwndProcess <> 0 Then
                    lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)


                    If lngReturn <> 0 Then
                        strModuleName = Space(MAX_PATH)
                        lngSize = 500
                        lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, lngSize)
                        strProcessName = Left(strModuleName, lngReturn)
                        strProcessName = UCase$(Trim$(strProcessName))
                        arrSplit = Split(strProcessName, "\")


                        If arrSplit(UBound(arrSplit)) = EXEName Then
                            GetProcessStatus = True
                            lngReturn = CloseHandle(lngHwndProcess)
                            Exit For
                        End If
                    End If
                End If
                lngReturn = CloseHandle(lngHwndProcess)


                DoEvents
                Next
End Function



