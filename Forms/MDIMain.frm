VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "ResourcesTutorial"
   ClientHeight    =   6930
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11265
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows-Standard
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbNew"
            Object.ToolTipText     =   "File New"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbOpen"
            Object.ToolTipText     =   "File Open"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbSave"
            Object.ToolTipText     =   "File Save"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      Begin VB.PictureBox PnlUserAuth 
         BorderStyle     =   0  'Kein
         Height          =   375
         Left            =   6120
         ScaleHeight     =   375
         ScaleWidth      =   5055
         TabIndex        =   2
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   0
            ScaleHeight     =   195
            ScaleWidth      =   2355
            TabIndex        =   6
            Top             =   30
            Width           =   2415
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "User:"
               Height          =   195
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Width           =   375
            End
            Begin VB.Label LblUserName 
               AutoSize        =   -1  'True
               Caption         =   "Username Default"
               Height          =   195
               Left            =   480
               TabIndex        =   7
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   2355
            TabIndex        =   3
            Top             =   30
            Width           =   2415
            Begin VB.Label LblAuthorName 
               AutoSize        =   -1  'True
               Caption         =   "Authorname Default"
               Height          =   195
               Left            =   600
               TabIndex        =   5
               Top             =   0
               Width           =   1410
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Author:"
               Height          =   195
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Width           =   510
            End
         End
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6630
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog FileDlg 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSep1 
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
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MDIChildren As Collection
Private m_FormArrange As FormArrangeConstants

Private Sub MDIForm_Load()
    If MDIChildren Is Nothing Then Set MDIChildren = New Collection
    LblUserName.Caption = Application.CurrentUserName
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
    Case QueryUnloadConstants.vbFormControlMenu
        'MsgBox "The user chose the Close command (the ""X"") from the Control menu on the form."
    Case QueryUnloadConstants.vbFormCode
        'MsgBox "The Unload statement is invoked from code."
    Case QueryUnloadConstants.vbAppWindows
        'MsgBox "The current Microsoft Windows operating environment session is ending."
    Case QueryUnloadConstants.vbAppTaskManager
        'MsgBox "The Microsoft Windows Task Manager is closing the application."
    Case QueryUnloadConstants.vbFormMDIForm
        'MsgBox "An MDI child form is closing because the MDI form is closing."
    Case QueryUnloadConstants.vbFormOwner
        'MsgBox "A form is closing because its owner is closing."
    End Select
    Dim frmChild 'As MDIChild
    For Each frmChild In Forms
        If frmChild Is MDIChild Then
            Cancel = frmChild.IsChanged
        End If
        If Cancel Then Exit Sub
    Next
End Sub

Private Sub MDIForm_Resize()
    Select Case m_FormArrange
    Case FormArrangeConstants.vbTileHorizontal, FormArrangeConstants.vbTileVertical: Me.Arrange m_FormArrange
    End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Application.Terminate
    Dim frm: For Each frm In Forms: Unload frm: Next
End Sub

Private Sub mnuFileNew_Click()
    Me.FileNew
End Sub

Public Sub mnuFileOpen_Click()
    Dim PFN As String
    If Application.OpenFileName_ShowDlg(PFN) = vbCancel Then Exit Sub
    FileOpen PFN
End Sub

Public Function FileNew() As Boolean
    Dim mc As MDIChild
    Set mc = New MDIChild
    MDIChildren.Add mc
    Dim dfn As String: dfn = Application.DefaultFileName
    mc.Caption = dfn & MDIChildren.Count & Application.MyExt
    mc.Show
End Function

Public Function FileOpen(PFN As String) As Boolean
    Dim mc As MDIChild
    If MDIChildren Is Nothing Then Set MDIChildren = New Collection
    If MDIChildren.Count > 0 Then
        Set mc = MDIChildren.Item(MDIChildren.Count)
    End If
    If mc Is Nothing Then
        Set mc = New MDIChild
        MDIChildren.Add mc
    End If
    If Len(mc.PathFileName) Or mc.IsChanged Then
        Set mc = New MDIChild
        MDIChildren.Add mc
    End If
    mc.Show
    mc.FileOpen PFN
End Function

Public Function FileClose(Child As MDIChild) As Boolean
    If Child.IsChanged Then
        Dim mr As VbMsgBoxResult: mr = MsgBox("Content is changed, do you want to save?", vbYesNoCancel)
        If mr = vbYes Then
            FileClose = Child.FileSave
            If Not FileClose Then Exit Function
        ElseIf mr = vbCancel Then
            Exit Function
        End If
    End If
    Dim c As MDIChild
    Dim i As Long, bfound As Boolean
    For i = 1 To MDIChildren.Count
        bfound = c Is Child
        If bfound Then Exit For
    Next
    If bfound Then MDIChildren.Remove i
    FileClose = True
    Unload Child
End Function

Public Sub mnuFileExit_Click()
    Unload Me
End Sub

Public Sub mnuExtrasOptions_Click()
    frmOptions.Show vbModal, MDIMain
End Sub

Public Sub WindowTileVertical()
    ArrangeChildForms FormArrangeConstants.vbTileVertical
End Sub
Public Sub WindowTileHorizontal()
    ArrangeChildForms FormArrangeConstants.vbTileHorizontal
End Sub
Public Sub WindowCascade()
    ArrangeChildForms FormArrangeConstants.vbCascade
End Sub
Public Sub WindowArrangeIcons()
    ArrangeChildForms FormArrangeConstants.vbArrangeIcons
End Sub
Private Sub ArrangeChildForms(fa As FormArrangeConstants)
    m_FormArrange = fa
    Me.Arrange m_FormArrange
End Sub

Public Sub mnuHelpInfo_Click()
    frmAbout.Show vbModal, MDIMain
End Sub


Public Sub LoadTBPics()
    'ImageList1.ListImages.Add ,,LoadResPicture(
End Sub
