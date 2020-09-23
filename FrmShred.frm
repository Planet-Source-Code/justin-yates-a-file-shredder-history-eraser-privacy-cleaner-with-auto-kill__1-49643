VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmShred 
   BorderStyle     =   0  'None
   ClientHeight    =   6255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7875
   ControlBox      =   0   'False
   Icon            =   "FrmShred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      Picture         =   "FrmShred.frx":0BC2
      ScaleHeight     =   6255
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox CheStart 
         BackColor       =   &H005A5D5B&
         Caption         =   "Run at Windows startup"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFC0&
         Height          =   2955
         Hidden          =   -1  'True
         Left            =   3345
         System          =   -1  'True
         TabIndex        =   4
         Top             =   525
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   1560
         Picture         =   "FrmShred.frx":39E5
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Timer TClosed 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2400
         Top             =   3600
      End
      Begin VB.Timer TCheck 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2040
         Top             =   3600
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFC0&
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   3135
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3135
      End
      Begin MSComctlLib.ListView lvCache 
         Height          =   1875
         Left            =   105
         TabIndex        =   6
         Top             =   4320
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777152
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "FILENAME"
            Text            =   "File Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "URL"
            Text            =   "Source URL"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Internet Cache"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Image BTempNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":59D7
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Image BTempDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":6BE9
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Image BTempActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":7DFB
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Image BFindNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":900D
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Image BFindDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":A21F
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Image BFindActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":B431
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Image BFolderNormal 
         Height          =   300
         Left            =   4980
         Picture         =   "FrmShred.frx":C643
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BFolderDown 
         Height          =   300
         Left            =   4980
         Picture         =   "FrmShred.frx":D855
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BFolderActive 
         Height          =   300
         Left            =   4980
         Picture         =   "FrmShred.frx":EA67
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BFileNormal 
         Height          =   300
         Left            =   3720
         Picture         =   "FrmShred.frx":FC79
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BFileDown 
         Height          =   300
         Left            =   3720
         Picture         =   "FrmShred.frx":10E8B
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BFileActive 
         Height          =   300
         Left            =   3720
         Picture         =   "FrmShred.frx":1209D
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BRunNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":132AF
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Image BRunDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":144C1
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Image BRunActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":156D3
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Image BRefreshNormal 
         Height          =   300
         Left            =   120
         Picture         =   "FrmShred.frx":168E5
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BRefreshDown 
         Height          =   300
         Left            =   120
         Picture         =   "FrmShred.frx":17AF7
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BRefreshActive 
         Height          =   300
         Left            =   120
         Picture         =   "FrmShred.frx":18D09
         ToolTipText     =   "Refresh History list"
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BRecentNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":19F1B
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Image BRecentDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":1B12D
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Image BRecentActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":1C33F
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Image BReadNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":1D551
         Top             =   600
         Width           =   1140
      End
      Begin VB.Image BReadDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":1E763
         Top             =   600
         Width           =   1140
      End
      Begin VB.Image BReadActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":1F975
         Top             =   600
         Width           =   1140
      End
      Begin VB.Image BPaintNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":20B87
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Image BPaintDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":21D99
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Image BPaintActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":22FAB
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Image BMediaNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":241BD
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Image BMediaDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":253CF
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Image BMediaActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":265E1
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Image BHistNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":277F3
         Top             =   960
         Width           =   1140
      End
      Begin VB.Image BHistDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":28A05
         Top             =   960
         Width           =   1140
      End
      Begin VB.Image BHistActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":29C17
         Top             =   960
         Width           =   1140
      End
      Begin VB.Image BCleanNormal 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":2AE29
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BCleanDown 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":2C03B
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BCleanActive 
         Height          =   300
         Left            =   6600
         Picture         =   "FrmShred.frx":2D24D
         Top             =   3660
         Width           =   1140
      End
      Begin VB.Image BMNormal 
         Height          =   300
         Left            =   7080
         Picture         =   "FrmShred.frx":2E45F
         Top             =   120
         Width           =   300
      End
      Begin VB.Image BMDown 
         Height          =   300
         Left            =   7080
         Picture         =   "FrmShred.frx":2E951
         Top             =   120
         Width           =   300
      End
      Begin VB.Image BMActive 
         Height          =   300
         Left            =   7080
         Picture         =   "FrmShred.frx":2EE43
         Top             =   120
         Width           =   300
      End
      Begin VB.Image BXNormal 
         Height          =   300
         Left            =   7440
         Picture         =   "FrmShred.frx":2F335
         Top             =   120
         Width           =   300
      End
      Begin VB.Image BXDown 
         Height          =   300
         Left            =   7440
         Picture         =   "FrmShred.frx":2F827
         Top             =   120
         Width           =   300
      End
      Begin VB.Image BXActive 
         Height          =   300
         Left            =   7440
         Picture         =   "FrmShred.frx":2FD19
         Top             =   120
         Width           =   300
      End
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu MnuHist 
         Caption         =   "Erase History"
      End
      Begin VB.Menu MnuAll 
         Caption         =   "Erase All"
      End
      Begin VB.Menu MnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmShred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
' Author:   Justin Yates                                            *
' VB ver:   VB6 SP5                                                 *
' Info:     Program to clear history, temp internet cache, cookies, *
'           media, paint, recent docs, search, find, run, & temp.   *
'           Also shreds files and has automatic cleaning after      *
'           closing internet explorer.
'********************************************************************
' Thanks to:                                                        *
' 1 John Bridle                                                     *
' 2 Anthony Christianson                                            *
' 3 Eduardo A. Morcillo                                             *
' http://www.mvps.org/emorcillo/en/index.shtml                      *
'********************************************************************
' The three index.dat files are loacted at:                         *
' Cookies\index.dat                                                 *
' History\History.IE5\index.dat                                     *
' Temporary Internet Files\Content.IE5\index.dat                    *
'********************************************************************

Option Explicit

Private HistoryPath As String
Private TempInternet As String
Private Cookies As String
Private TempPath As String
Private fso As New FileSystemObject
Private F As File
Private FD As Folder
Private FLS As Files
Private FDS As Folders

Dim response As Long
Dim FirstOpen As Boolean
Dim ExplorerClosed As Boolean
Dim History As CURLHistory

Private Const ChunkSize As Integer = 8192

' api declarations
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub CleanAll()
    KillHistory
    CleanFind
    CleanMediaPlayer
    CleanPaintBrushFileLst
    CleanDocs
    CleanRunMnu
    KillTemp
End Sub

' delete folder and it's contents
Private Sub ShredFolder()
    On Error Resume Next
    MousePointer = vbHourglass
    response = MsgBox("Shred " & Dir1.Path & " and all its contents?", vbYesNo + vbExclamation, Me.Caption)
    If response = vbYes Then
        KillFiles (Dir1.Path)
        fso.DeleteFolder Dir1.Path, True
        Drive1_Change
    End If
    MousePointer = vbNormal
End Sub

' kill history
Private Sub KillHistory()
    ' erase all the cookies
    Set FD = fso.GetFolder(Cookies)
    Set FLS = FD.Files
    For Each F In FLS
        If Right(F.Path, 9) = "index.dat" Then
            ' delete index.dat from Cookies
            DeleteFile F.Path
        Else
            ShredFile F.Path
        End If
    Next
    
    FindCacheEntries ' first we load the cache
    ' loop trough the entries...
    Do While lvCache.ListItems.Count > 0
        ' ...and delete them
        DeleteCacheEntry lvCache.ListItems(1).SubItems(1)
        lvCache.ListItems.Remove 1
    Loop
    ' delete index.dat from Temporary Internet files
    DeleteFile (NormalisePath(TempInternet) & "Content.IE5\index.dat")
    
    CleanIE ' Clear Typed URLs
    DeleteFile (NormalisePath(HistoryPath) & "history.ie5\index.dat") ' delete index.dat from History
    History.Clear ' clear History
    History.Refresh ""
End Sub

' Shred a file
Private Sub ShredTheFile()
    On Error GoTo ErrHandler
    Dim FilePath As String
    MousePointer = vbHourglass
    FilePath = NormalisePath(File1.Path) & File1.FileName
    response = MsgBox("Shred " & FilePath & " ?", vbYesNo + vbQuestion, Me.Caption)
    If response = vbYes Then
        ShredFile (FilePath)
        File1.Refresh
    End If
    MousePointer = vbNormal
    Exit Sub
ErrHandler:
    Close
    MousePointer = vbNormal
    MsgBox Err.Description, vbOKOnly + vbExclamation, Me.Caption
End Sub
' Clear temp directory
Private Sub KillTemp()
    MousePointer = vbHourglass
    KillFiles (TempPath)
    Set FD = fso.GetFolder(TempPath)
    Set FDS = FD.SubFolders
    For Each FD In FDS
        fso.DeleteFolder FD, True
    Next
    MousePointer = vbNormal
End Sub

Private Sub KillFiles(MyPath As String)
    On Error Resume Next
    Set FD = fso.GetFolder(MyPath)
    Set FLS = FD.Files
    For Each F In FLS ' first delete the files in the folder
        ShredFile (F.Path)
    Next
    Set FDS = FD.SubFolders
    For Each FD In FDS ' now delete the files in every folder
        KillFiles (FD.Path) ' using recursion
    Next
End Sub
'*******************************************************************************
' the next few subs are for the buttons to have effects: active, down and normal
Private Sub BCleanActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BCleanActive.Visible = False
    BCleanDown.Visible = True
End Sub

Private Sub BCleanActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BCleanDown.Visible = False
    BCleanNormal.Visible = True
    CleanAll
End Sub

Private Sub BCleanNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BCleanActive.Visible = True
    BCleanDown.Visible = False
    BCleanNormal.Visible = False
End Sub

Private Sub BFileActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFileActive.Visible = False
    BFileDown.Visible = True
End Sub

Private Sub BFileActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFileDown.Visible = False
    BFileNormal.Visible = True
    ShredTheFile
End Sub

Private Sub BFileNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFileActive.Visible = True
    BFileDown.Visible = False
    BFileNormal.Visible = False
End Sub

Private Sub BFindActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFindActive.Visible = False
    BFindDown.Visible = True
End Sub

Private Sub BFindActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFindDown.Visible = False
    BFindNormal.Visible = True
    CleanFind ' clean find recent items
End Sub

Private Sub BFindNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFindActive.Visible = True
    BFindDown.Visible = False
    BFindNormal.Visible = False
End Sub

Private Sub BFolderActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFolderActive.Visible = False
    BFolderDown.Visible = True
End Sub

Private Sub BFolderActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFolderDown.Visible = False
    BFolderNormal.Visible = True
    ShredFolder ' shred a folder and all its contents
End Sub

Private Sub BFolderNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BFolderActive.Visible = True
    BFolderDown.Visible = False
    BFolderNormal.Visible = False
End Sub

Private Sub BHistActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BHistActive.Visible = False
    BHistDown.Visible = True
End Sub

Private Sub BHistActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BHistDown.Visible = False
    BHistNormal.Visible = True
    KillHistory
End Sub

Private Sub BHistNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BHistActive.Visible = True
    BHistDown.Visible = False
    BHistNormal.Visible = False
End Sub

Private Sub BMActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMActive.Visible = False
    BMDown.Visible = True
End Sub

Private Sub BMActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMDown.Visible = False
    BMNormal.Visible = True
    Me.Hide
End Sub

Private Sub BMediaActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMediaActive.Visible = False
    BMediaDown.Visible = True
End Sub

Private Sub BMediaActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMediaDown.Visible = False
    BMediaNormal.Visible = True
     CleanMediaPlayer ' clear windows media players recent file list
End Sub

Private Sub BMediaNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMediaActive.Visible = True
    BMediaDown.Visible = False
    BMediaNormal.Visible = False
End Sub

Private Sub BMNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BMActive.Visible = True
    BMDown.Visible = False
    BMNormal.Visible = False
End Sub

Private Sub BPaintActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BPaintActive.Visible = False
    BPaintDown.Visible = True
End Sub

Private Sub BPaintActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BPaintDown.Visible = False
    BPaintNormal.Visible = True
    CleanPaintBrushFileLst ' clear Paint's recent file list
End Sub

Private Sub BPaintNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BPaintActive.Visible = True
    BPaintDown.Visible = False
    BPaintNormal.Visible = False
End Sub

Private Sub BReadActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BReadActive.Visible = False
    BReadDown.Visible = True
End Sub

Private Sub BReadActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BReadDown.Visible = False
    BReadNormal.Visible = True
    FrmAbout.Show vbModal
End Sub

Private Sub BReadNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BReadActive.Visible = True
    BReadDown.Visible = False
    BReadNormal.Visible = False
End Sub

Private Sub BRecentActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRecentActive.Visible = False
    BRecentDown.Visible = True
End Sub

Private Sub BRecentActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRecentDown.Visible = False
    BRecentNormal.Visible = True
    CleanDocs ' clean recent documents
End Sub

Private Sub BRecentNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRecentActive.Visible = True
    BRecentDown.Visible = False
    BRecentNormal.Visible = False
End Sub

Private Sub BRefreshActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRefreshActive.Visible = False
    BRefreshDown.Visible = True
End Sub

Private Sub BRefreshActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRefreshDown.Visible = False
    BRefreshNormal.Visible = True
    FindCacheEntries
End Sub

Private Sub BRefreshNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRefreshActive.Visible = True
    BRefreshDown.Visible = False
    BRefreshNormal.Visible = False
End Sub

Private Sub BRunActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRunActive.Visible = False
    BRunDown.Visible = True
End Sub

Private Sub BRunActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRunDown.Visible = False
    BRunNormal.Visible = True
    CleanRunMnu ' cear run dialog box (must restart computer for settings to take effect)
End Sub

Private Sub BRunNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BRunActive.Visible = True
    BRunDown.Visible = False
    BRunNormal.Visible = False
End Sub

Private Sub BTempActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BTempActive.Visible = False
    BTempDown.Visible = True
End Sub

Private Sub BTempActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BTempDown.Visible = False
    BTempNormal.Visible = True
    KillTemp ' clear the contents of your temp folder
End Sub

Private Sub BTempNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BTempActive.Visible = True
    BTempDown.Visible = False
    BTempNormal.Visible = False
End Sub

Private Sub BXActive_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    BXActive.Visible = False
    BXDown.Visible = True
End Sub

Private Sub BXActive_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    BXDown.Visible = False
    BXNormal.Visible = True
    MnuExit_Click
End Sub

Private Sub BXNormal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BXActive.Visible = True
    BXDown.Visible = False
    BXNormal.Visible = False
End Sub

Private Sub CheStart_Click()
    If CheStart.Value = 1 Then
        savestring HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Shredder", NormalisePath(App.Path) & "Shredder.exe -min"
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Shredder"
    End If
End Sub

' end of button effects
'*******************************************************************************
Private Sub Dir1_Change()
    File1.Path = Dir1.Path ' set the filelist path when a directory is changed
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrDrive
    Dir1.Path = Drive1.Drive
    Exit Sub
ErrDrive:
    MsgBox Err.Description, vbOKOnly + vbExclamation, Me.Caption
End Sub

Private Sub Form_Load()
    Dim StartUp As String
    ShowIcon Me, "Shredder", Me.Icon.Handle ' put the icon in the system tray
    If Command = "-min" Then
        Me.Hide
    End If
        
    FindCacheEntries ' load cache
    
    Set History = New CURLHistory ' create an instance
    
    ' get the required paths form registry
    HistoryPath = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "History") ' history folder
    TempInternet = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache") ' temporary internet files folder
    Cookies = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cookies") ' cookies folder
    TempPath = fso.GetSpecialFolder(2) ' windows temp folder
    
    StartUp = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Shredder")
    If StartUp = vbNullString Then
        CheStart.Value = 0
    Else
        CheStart.Value = 1
    End If
    
    Dir1.Path = App.Path
    TCheck.Enabled = True
End Sub

Private Sub FindCacheEntries()
    
    lvCache.ListItems.Clear
    
    ' Find first cache item
    If FindFirstCacheEntry() Then
        ' If File name nothing then add the source url
        If CachedEntryFileName = vbNullString Then
            lvCache.ListItems.Add , , CachedEntrySourceURL
        Else
            lvCache.ListItems.Add , , CachedEntryFileName
        End If
        ' Add the source url
        lvCache.ListItems(lvCache.ListItems.Count).SubItems(1) = CachedEntrySourceURL
                
        ' Loop until there are no more cache items
        Do While FindNextCacheEntry
            If CachedEntryCacheType And &H1 Then
                ' If File name nothing then add the source url
                If CachedEntryFileName = vbNullString Then
                    lvCache.ListItems.Add , , CachedEntrySourceURL
                Else
                    lvCache.ListItems.Add , , CachedEntryFileName
                End If
                ' Add the source url
                lvCache.ListItems(lvCache.ListItems.Count).SubItems(1) = CachedEntrySourceURL
            End If
        Loop
    End If
    
    ReleaseCache
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
    
        Case LeftDblClick
            FindCacheEntries
            Me.Show
        Case RightMouseDown
            PopupMenu Mnu_File, 0 ' show popup menu from system tray

    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HideIcon ' remove the icon from the system tray
End Sub

' erase all
Private Sub MnuAll_Click()
    CleanAll
End Sub

Private Sub MnuExit_Click()
    HideIcon ' remove the icon from the system tray
    End
End Sub

Private Sub MnuHist_Click()
    KillHistory
End Sub

Private Sub MnuShow_Click()
    FindCacheEntries
    Me.Show
End Sub

' Thanx to John for his help in optimizing this process
Private Sub ShredFile(Path As String)
    On Error Resume Next
    Dim i As Integer
    Dim j As Long
    Dim fnum As Long
    Dim fLen As Long
    Dim idx As Long
    Dim str As String
    
    fnum = FreeFile
    ' IMPORTANT DONT CLEAR FILE FIRST...OVERWRITE EXISTING
    ' YOU COULD START WRITING TO NEW DISK SECTOR!!!
    Open Path For Input As fnum
        fLen = LOF(fnum)
    Close fnum
    fLen = fLen / ChunkSize
    
    SetAttr Path, vbNormal ' to remove any read only attributes
    str = String$(ChunkSize, Chr(35)) 'A pre-Buffered String with only one character
                                  'This speeds up the overwriting process tenfold
                                  'and it doesnt matter if you variate the character
                                  'overwritten is overwritten and you only cause more
                                  'overheads...funny works quicker with Chr(35) than "#"
    
    For i = 0 To 10
        fnum = FreeFile
        Open Path For Binary As fnum
        ' fill the file with rubbish
        For j = 0 To fLen
            'Create the start idx of where to continue the file
            idx = (ChunkSize * j) + 1
            Put fnum, idx, str
            DoEvents
        Next j
        Close fnum
        DoEvents
    Next i
    DeleteFile (Path)
End Sub

' reset the buttons to normal
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    BXActive.Visible = False
    BXDown.Visible = False
    BXNormal.Visible = True
    BMActive.Visible = False
    BMDown.Visible = False
    BMNormal.Visible = True
    BCleanActive.Visible = False
    BCleanDown.Visible = False
    BCleanNormal.Visible = True
    BHistActive.Visible = False
    BHistDown.Visible = False
    BHistNormal.Visible = True
    BMediaActive.Visible = False
    BMediaDown.Visible = False
    BMediaNormal.Visible = True
    BPaintActive.Visible = False
    BPaintDown.Visible = False
    BPaintNormal.Visible = True
    BReadActive.Visible = False
    BReadDown.Visible = False
    BReadNormal.Visible = True
    BRecentActive.Visible = False
    BRecentDown.Visible = False
    BRecentNormal.Visible = True
    BRefreshActive.Visible = False
    BRefreshDown.Visible = False
    BRefreshNormal.Visible = True
    BRunActive.Visible = False
    BRunDown.Visible = False
    BRunNormal.Visible = True
    BFileActive.Visible = False
    BFileDown.Visible = False
    BFileNormal.Visible = True
    BFolderActive.Visible = False
    BFolderDown.Visible = False
    BFolderNormal.Visible = True
    BFindActive.Visible = False
    BFindDown.Visible = False
    BFindNormal.Visible = True
    BTempActive.Visible = False
    BTempDown.Visible = False
    BTempNormal.Visible = True
    DoEvents
    Dim lngReturnValue As Long
    ' move the form with mouse click any where
    If Button = 1 Then
        'Release capture
        Call ReleaseCapture
        'Send a 'left mouse button down on caption'-message to our form
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

' check to see if Internet explorer is open, so we can auto kill when it closes
Private Sub TCheck_Timer()
    FirstOpen = GetProcessStatus("iexplore.exe") ' check the task manager for "iexplore.exe" process
    If FirstOpen = True Then
        TCheck.Enabled = False
        TClosed.Enabled = True ' now monitor "iexplore.exe"
    End If
End Sub

' if internet explorer has been opened and all instance have been close auto kill history
Private Sub TClosed_Timer()
    ExplorerClosed = GetProcessStatus("iexplore.exe")
    If ExplorerClosed = False Then
        nfIconData.hIcon = Picture1.Picture.Handle ' change the system tray icon to show user history is been auto killed
        Call Shell_NotifyIcon(NIF_MESSAGE, nfIconData)
        Sleep 500 ' to allow the icon to been seen, otherwise its too fast to see
        KillHistory
        nfIconData.hIcon = Me.Icon.Handle
        Call Shell_NotifyIcon(NIF_MESSAGE, nfIconData) ' reset the icon in the system tray back to original
        TClosed.Enabled = False
        TCheck.Enabled = True ' start checking for first instance of internet explorer again
    End If
End Sub
