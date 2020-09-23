VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "URLHistoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Dim m_STAT As STATURL
'*********************************************************************************************
'
' Returns a string from a unicode string pointer.
'
'*********************************************************************************************
Private Function Ptr2StrW(Ptr As Long) As String
Dim sRtn() As Byte

    ' Check if the pointer is valid
    If Ptr <> 0 Then
    
        ReDim sRtn(lstrlenW(ByVal Ptr) * 2)
        
        ' Copy the string to the byte array
        MoveMemory sRtn(0), ByVal Ptr, UBound(sRtn)
        Ptr2StrW = sRtn()
        
    End If
    
End Function

'*********************************************************************************************
'
' Initializes the object data
'
'*********************************************************************************************
Friend Sub Init(STAT As STATURL)

    m_STAT = STAT
    
End Sub

'*********************************************************************************************
'
' Returns the URL
'
'*********************************************************************************************
Public Property Get URL() As String

    URL = Ptr2StrW(m_STAT.pwcsUrl)
    
End Property

'*********************************************************************************************
'
' Returns the URL title
'
'*********************************************************************************************
Public Property Get Title() As String

    Title = Ptr2StrW(m_STAT.pwcsTitle)
    
    If Title = "" Then Title = "     (No Title Given)"
        
End Property

Private Sub Class_Terminate()

    ' Release the strings from
    ' m_STAT variable
    CoTaskMemFree m_STAT.pwcsTitle
    CoTaskMemFree m_STAT.pwcsUrl
    
End Sub

