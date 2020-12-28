VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Settings"
   ClientHeight    =   4575
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   14295
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox tabPage1 
      BorderStyle     =   0  'Kein
      Height          =   3300
      Left            =   240
      ScaleHeight     =   3300
      ScaleWidth      =   5445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5445
      Begin VB.CommandButton BtnDeleteAllSettings 
         Caption         =   "Delete All Settings"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton BtnRegisterShellFileTypes 
         Caption         =   "Register Shell File Type"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton BtnUnRegisterShellFileTypes 
         Caption         =   "Unregister Shell File Type"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Un-/Register File Type"
            Key             =   "Group1"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Startup"
            Key             =   "Group2"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Language"
            Key             =   "Group3"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox tabPage3 
      BorderStyle     =   0  'Kein
      Height          =   2820
      Left            =   10320
      ScaleHeight     =   2820
      ScaleWidth      =   3765
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   3765
      Begin VB.PictureBox pbFlag4 
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   1800
         Width           =   615
      End
      Begin VB.PictureBox pbFlag3 
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox pbFlag2 
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox pbFlag1 
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Greek"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hrvatsk"
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "German"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "English"
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox tabPage2 
      BorderStyle     =   0  'Kein
      Height          =   2820
      Left            =   6360
      ScaleHeight     =   2820
      ScaleWidth      =   3645
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   3645
      Begin VB.CheckBox ChkShowLoginAtStartup 
         Caption         =   "Show Login at startup"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox ChkShowTippsAtStartup 
         Caption         =   "Show Tipps at startup"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox ChkShowSplashAtStartup 
         Caption         =   "Show Splashscreen at startup"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.CommandButton BtnApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Form_Load()
    
    Dim bkColor As Long: bkColor = GetBkColor(GetDC(tbsOptions.hWnd))
    Debug.Print Hex(bkColor)
    BackgroundColorAndAllChildren(tabPage1) = bkColor
    BackgroundColorAndAllChildren(tabPage2) = bkColor
    BackgroundColorAndAllChildren(tabPage3) = bkColor
    
    'Me.ScaleWidth = 6255
    Me.Width = 6345
    ChkShowSplashAtStartup.Value = Abs(Settings.ShowSplashAtStartup)
    ChkShowTippsAtStartup.Value = Abs(Settings.ShowTippsAtStartup)
    ChkShowLoginAtStartup.Value = Abs(Settings.ShowLoginAtStartup)
        
    Select Case Settings.Language
    Case 1: Option1.Value = True
    Case 2: Option2.Value = True
    Case 3: Option3.Value = True
    Case 4: Option4.Value = True
    End Select
    
    BtnRegisterShellFileTypes.Caption = "Register Shell File Type: *" & Application.MyExt
    
    'center Form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub
Private Property Let BackgroundColorAndAllChildren(ctrl As PictureBox, ByVal Color As Long)
    ctrl.BackColor = Color
    Dim c
    For Each c In Me.Controls
        If c.Container Is ctrl Then
            c.BackColor = Color
        End If
    Next
End Property
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'STRG+TAB, to move to the next tabpage
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tabpage -> move to the first
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub BtnDeleteAllSettings_Click()
    Settings.Delete
End Sub

Private Sub BtnRegisterShellFileTypes_Click()
    Application.RegisterExt
End Sub

Private Sub BtnUnRegisterShellFileTypes_Click()
    Application.UnRegisterExt
End Sub

Private Sub BtnApply_Click()
    Settings.ShowSplashAtStartup = ChkShowSplashAtStartup.Value = vbChecked
    Settings.ShowTippsAtStartup = ChkShowTippsAtStartup.Value = vbChecked
    Settings.ShowLoginAtStartup = ChkShowLoginAtStartup.Value = vbChecked
    Dim el As ELanguage
    Select Case True
    Case Option1.Value: el = 1
    Case Option2.Value: el = 2
    Case Option3.Value: el = 3
    Case Option4.Value: el = 4
    End Select
    Settings.Language = el
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnOK_Click()
    BtnApply_Click
    Unload Me
End Sub

Private Sub tbsOptions_Click()
    
    Select Case tbsOptions.SelectedItem.Index
    Case 1: tabPage1.ZOrder 0
    Case 2: tabPage2.ZOrder 0: tabPage2.Move tabPage1.Left, tabPage1.Top, tabPage1.Width, tabPage1.Height
    Case 3: tabPage3.ZOrder 0: tabPage3.Move tabPage1.Left, tabPage1.Top, tabPage1.Width, tabPage1.Height
    Case Else
    End Select
    
End Sub
