VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About File Shredder"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4215
      Begin VB.Label lblSource 
         Caption         =   "View Source"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         MouseIcon       =   "FrmAbout.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Tag             =   "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49643&lngWId=1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "The source code can be viewed below"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "This program is freeware and comes with no guarantees."
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   "The author of this program will not be held responsible for any damage that can be caused by the use of this program."
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmAbout.frx":074C
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Author: Justin Yates"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1680
         Picture         =   "FrmAbout.frx":0858
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Hyperlink(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblSource
        .FontUnderline = False
    End With
End Sub

Private Sub lblSource_Click()
    Hyperlink lblSource.Tag
End Sub

Private Sub lblSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblSource
        .FontUnderline = True
    End With
End Sub
