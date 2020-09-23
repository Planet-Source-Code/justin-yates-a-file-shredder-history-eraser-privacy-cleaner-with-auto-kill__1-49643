Attribute VB_Name = "modCache"
Option Explicit

Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40

Private Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize As Long
   lpszSourceUrlName As Long
   lpszLocalFileName As Long
   CacheEntryType As Long
   dwUseCount As Long
   dwHitRate As Long
   dwSizeLow As Long
   dwSizeHigh As Long
   LastModifiedTime As FILETIME
   ExpireTime As FILETIME
   LastAccessTime As FILETIME
   LastSyncTime As FILETIME
   lpHeaderInfo As Long
   dwHeaderInfoSize As Long
   lpszFileExtension As Long
   dwExemptDelta As Long
End Type

Private hEnumHandle As Long
Private ci As INTERNET_CACHE_ENTRY_INFO
Private lPtrCI As Long

' Functions
Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszSearchPattern As String, ByVal lpCacheInfo As Long, lpdwFirstCacheEntryInfoBufferSize As Long) As Long

Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, ByVal lpCacheInfo As Long, lpdwNextCacheEntryInfoBufferSize As Long) As Long

Private Declare Function FindCloseUrlCache Lib "wininet.dll" (ByVal hEnumHandle As Long) As Long

Private Declare Function GetUrlCacheEntryInfo Lib "wininet.dll" Alias "GetUrlCacheEntryInfoA" (ByVal lpszUrlName As String, ByVal lpCacheInfo As Long, lpdwCacheEntryInfoBufferSize As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
        
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Long, pSource As Long, ByVal dwLength As Long)
        
Private Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" (pDest As INTERNET_CACHE_ENTRY_INFO, pSource As Long, ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
    
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Long) As Long

Public Function CachedEntryCacheType() As Long
    CachedEntryCacheType = ci.CacheEntryType
    
End Function

Public Function CachedEntrySourceURL() As String
    Dim strdata As String
    Dim lReturnValue As Long
    Dim iPosition As Long
    
    ' Allocate a buffer for our filename
    strdata = String$(lstrlenA(ci.lpszSourceUrlName), 0)
    
    ' Copy the data to our buffer
    lReturnValue = lstrcpyA(strdata, ci.lpszSourceUrlName)
    
    ' If successful then get the data we need
    If lReturnValue Then
        CachedEntrySourceURL = strdata
    End If

End Function

Public Function DeleteCacheEntry(SourceUrl As String) As Boolean
    Dim lReturnValue As Long
    
    lReturnValue = DeleteUrlCacheEntry(SourceUrl)
    DeleteCacheEntry = CBool(lReturnValue)
    
End Function
' This function searches the cache for the cache entry corresponding to the given url
Public Function FindEntryInCache(Url As String) As Boolean
    Dim lReturnValue As Long, lSizeOfStruct As Long
    
    ' Find out the required size for the item
    lReturnValue = GetUrlCacheEntryInfo(Url, 0&, lSizeOfStruct)

    ' If we have memory allocated, free it
    If lPtrCI Then
        LocalFree lPtrCI
    End If

    ' Allocate a buffer
    lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
    
    If lPtrCI Then
        ' Copy from the buffer
        CopyMemory ByVal lPtrCI, lSizeOfStruct, 4

        lReturnValue = GetUrlCacheEntryInfo(Url, lPtrCI, lSizeOfStruct)
        '' copy the memory that our pointer points to into our structure
        CopyMemory1 ci, ByVal lPtrCI, Len(ci)
        ' Free the allocated memory
        LocalFree lPtrCI
    End If
    
    ' Was it successful?
    FindEntryInCache = CBool(lReturnValue)
    
    
End Function

Public Function FindFirstCacheEntry() As Boolean
    Dim lSizeOfStruct As Long
    
    'The FindFirstUrlCacheEntry function returns a handle which can be used with subsequent calls to the FindNextUrlCacheEntry function
    
    ' First see if we have already opened a search handle, if so close it
    If hEnumHandle <> 0 Then
        FindCloseUrlCache hEnumHandle
    End If
        
    ' Find out the required size for the item
    hEnumHandle = FindFirstUrlCacheEntry(vbNullString, 0&, lSizeOfStruct)
    
    ' If we have memory allocated, free it
    If lPtrCI Then
        LocalFree lPtrCI
    End If

    ' Allocate a buffer
    lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
    
    If lPtrCI Then
        ' Copy from the buffer
        CopyMemory ByVal lPtrCI, lSizeOfStruct, 4
        
        hEnumHandle = FindFirstUrlCacheEntry(ByVal vbNullString, lPtrCI, lSizeOfStruct)
        
        ' Copy the memory that our pointer points to into our structure
        CopyMemory1 ci, ByVal lPtrCI, Len(ci)
        
    End If
    
    ' Was it successful?
    FindFirstCacheEntry = CBool(hEnumHandle)
    
End Function

Public Function FindNextCacheEntry() As Boolean
    Dim lReturnValue As Long, lSizeOfStruct As Long
    
    If hEnumHandle <> 0 Then
        ' Find out the required size for the next item
        lReturnValue = FindNextUrlCacheEntry(hEnumHandle, 0&, lSizeOfStruct)
                
        ' If we have memory allocated, free it
        If lPtrCI Then
            LocalFree lPtrCI
        End If

        ' Allocate a buffer
        lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
        
        If lPtrCI Then
            ' Copy from the buffer
            CopyMemory ByVal lPtrCI, lSizeOfStruct, 4
            
            lReturnValue = FindNextUrlCacheEntry(hEnumHandle, lPtrCI, lSizeOfStruct)
            ' Copy the memory that our pointer points to into our structure
            CopyMemory1 ci, ByVal lPtrCI, Len(ci)
        End If

        If lReturnValue <> 0 Then
            FindNextCacheEntry = CBool(lReturnValue)
        End If
        
    End If
    
End Function

Public Function CachedEntryFileName() As String
    Dim strdata As String
    Dim lReturnValue As Long
    Dim iPosition As Long
    
    ' Allocate a buffer for our filename
    strdata = String$(lstrlenA(ByVal ci.lpszLocalFileName), 0)
    
    ' Now copy the data to our buffer
    lReturnValue = lstrcpyA(strdata, ci.lpszLocalFileName)
    
    If lReturnValue Then
        CachedEntryFileName = strdata
    End If
    
End Function

Public Sub ReleaseCache()
  
    If hEnumHandle Then
        Call FindCloseUrlCache(hEnumHandle)
    End If
    
End Sub

