VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuka File Berdasarkan Ekstensi Program"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
      Const SW_SHOWNORMAL = 1
      Const SE_ERR_FNF = 2&
      Const SE_ERR_PNF = 3&
      Const SE_ERR_ACCESSDENIED = 5&
      Const SE_ERR_OOM = 8&
      Const SE_ERR_DLLNOTFOUND = 32&
      Const SE_ERR_SHARE = 26&
      Const SE_ERR_ASSOCINCOMPLETE = 27&
      Const SE_ERR_DDETIMEOUT = 28&
      Const SE_ERR_DDEFAIL = 29&
      Const SE_ERR_DDEBUSY = 30&
      Const SE_ERR_NOASSOC = 31&
      Const ERROR_BAD_FORMAT = 11&

Function OpenDocument(ByVal DocName As String) As Long
     Dim Scr_hDC As Long
    'Scr_hDC = GetDesktopWindow()
     OpenDocument = ShellExecute(Me.hwnd, "Open", _
     DocName, "", "C:\", SW_SHOWNORMAL)
      End Function

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub
    
Private Sub File1_DblClick()
Dim r As Long, msg As String
          Dim str As String
          If Right(Dir1.Path, 1) = "\" Then
              str = Dir1.Path & File1.FileName
          Else
              str = Dir1.Path & "\" & File1.FileName
          End If
          Me.Caption = str
          r = OpenDocument(str)
          If r <= 32 Then
              Select Case r
                  Case SE_ERR_FNF
                      msg = "File not found"
                  Case SE_ERR_PNF
                      msg = "Path not found"
                  Case SE_ERR_ACCESSDENIED
                      msg = "Access denied"
                  Case SE_ERR_OOM
                      msg = "Out of memory"
                  Case SE_ERR_DLLNOTFOUND
                      msg = "DLL not found"
                  Case SE_ERR_SHARE
                      msg = "A sharing violation occurred"
                  Case SE_ERR_ASSOCINCOMPLETE
                      msg = "Incomplete or invalid file association"
                  Case SE_ERR_DDETIMEOUT
                      msg = "DDE Time out"
                  Case SE_ERR_DDEFAIL
                      msg = "DDE transaction failed"
                  Case SE_ERR_DDEBUSY
                      msg = "DDE busy"
                  Case SE_ERR_NOASSOC
                      msg = "No association for file extension"
                  Case ERROR_BAD_FORMAT
                      msg = "Invalid EXE file or error in EXE image"
                  Case Else
                      msg = "Unknown error"
              End Select
              MsgBox msg
          End If
End Sub


