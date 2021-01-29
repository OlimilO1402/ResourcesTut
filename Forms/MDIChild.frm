VERSION 5.00
Begin VB.Form MDIChild 
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   9555
   Icon            =   "MDIChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   9555
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileclose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuExtrasOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowTileVert 
         Caption         =   "Tile &verticaly"
      End
      Begin VB.Menu mnuWindowTileHoriz 
         Caption         =   "Tile &horizontaly"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "MDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AuthorName As String
Public IsChanged As Boolean
Private m_PathFileName As String

Private Sub Form_Load()
    Me.AuthorName = Application.CurrentUserName
    Me.WindowState = Settings.MDIChildWindowState
    MDIMain.LblAuthorName.Caption = Me.AuthorName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not MDIMain.FileClose(Me)
'    Select Case UnloadMode
'    Case QueryUnloadConstants.vbFormControlMenu
'        'MsgBox "The user chose the Close command (the ""X"") from the Control menu on the form."
'    Case QueryUnloadConstants.vbFormCode
'        'MsgBox "The Unload statement is invoked from code."
'    Case QueryUnloadConstants.vbAppWindows
'        'MsgBox "The current Microsoft Windows operating environment session is ending."
'    Case QueryUnloadConstants.vbAppTaskManager
'        'MsgBox "The Microsoft Windows Task Manager is closing the application."
'    Case QueryUnloadConstants.vbFormMDIForm
'        'MsgBox "An MDI child form is closing because the MDI form is closing."
'    Case QueryUnloadConstants.vbFormOwner
'        'MsgBox "A form is closing because its owner is closing."
'    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Settings.MDIChildWindowState = Me.WindowState
End Sub
Private Sub Form_Resize()
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight
    If W > 0 And H > 0 Then Text1.Move 0, 0, W, H
End Sub

Public Property Let PathFileName(RHS As String)
    m_PathFileName = RHS
    Me.Caption = m_PathFileName
End Property
Public Property Get PathFileName() As String
    PathFileName = m_PathFileName
End Property

Public Function FileOpen(PFN As String) As Boolean
    Me.PathFileName = PFN
    Dim Content As String
    If ReadFile(PFN, Content) Then
        Text1.Text = Content
        FileOpen = True
        Exit Function
    End If
    'sonst Fehler
    MsgBox "File access denied or file not found: " & vbCrLf & PFN
End Function

Public Function FileSave() As Boolean
    If Len(m_PathFileName) = 0 Then
        FileSave = FileSaveAs
    Else
        If WriteFile(Me.PathFileName, Text1.Text) Then
            FileSave = True
            Me.IsChanged = False
        End If
    End If
End Function
Public Function FileSaveAs() As Boolean
    Dim PFN As String
    If Application.SaveFileName_ShowDlg(PFN) = vbCancel Then Exit Function
    Me.PathFileName = PFN
    If WriteFile(PFN, Text1.Text) Then
        FileSaveAs = True
        Me.IsChanged = False
    End If
End Function

Private Sub mnuFileNew_Click()
    MDIMain.FileNew
End Sub
Private Sub mnuFileOpen_Click()
    MDIMain.mnuFileOpen_Click
End Sub
Private Sub mnuFileClose_Click()
    MDIMain.FileClose Me
End Sub
Private Sub mnuFileSave_Click()
    Me.FileSave
End Sub
Private Sub mnuFileSaveAs_Click()
    Me.FileSaveAs
End Sub

Private Sub mnuFileExit_Click()
    MDIMain.mnuFileExit_Click
End Sub

Private Sub mnuExtrasOptions_Click()
    MDIMain.mnuExtrasOptions_Click
End Sub

Private Sub mnuWindowTileVert_Click()
    MDIMain.WindowTileVertical
End Sub
Private Sub mnuWindowTileHoriz_Click()
    MDIMain.WindowTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
    MDIMain.WindowCascade
End Sub
Private Sub mnuWindowArrangeIcons_Click()
    MDIMain.WindowArrangeIcons
End Sub

Private Sub mnuHelpInfo_Click()
    MDIMain.mnuHelpInfo_Click
End Sub

Private Sub Text1_Change()
    IsChanged = True
End Sub

Private Function WriteFile(PFN As String, Content As String) As Boolean
Try: On Error GoTo Finally
    If FileExists(PFN) Then Kill PFN
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , Content
    WriteFile = True
Finally:
    Close FNr
End Function
    
Private Function ReadFile(PFN As String, Content_out As String) As Boolean
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Read As FNr
    Content_out = Space(LOF(FNr))
    Get FNr, , Content_out
    ReadFile = True
Finally:
    Close FNr
End Function

Private Function FileExists(PFN As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(PFN) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Private Function PathExists(ByVal DirectoryName As String) As Boolean
    On Error Resume Next
    PathExists = CBool(GetAttr(DirectoryName) And vbDirectory)
    On Error GoTo 0
End Function

