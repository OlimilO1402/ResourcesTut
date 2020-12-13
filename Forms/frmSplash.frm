VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   360
      End
      Begin VB.Label lblCountDown 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company: Users company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning: copying is herby allowed and granted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   6
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Rechts
         Caption         =   "Licensed for you"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Company/Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   6825
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   600
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CountDown As Integer

Private Sub Form_Load()
    CountDown = 5
    lblCompanyProduct.Caption = App.CompanyName
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblPlatform.Caption = "Windows"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer): Unload Me: End Sub
Private Sub Frame1_Click():                     Unload Me: End Sub
Private Sub imgLogo_Click():                    Unload Me: End Sub
Private Sub lblCompany_Click():                 Unload Me: End Sub
Private Sub lblCompanyProduct_Click():          Unload Me: End Sub
Private Sub lblCopyright_Click():               Unload Me: End Sub
Private Sub lblLicenseTo_Click():               Unload Me: End Sub
Private Sub lblPlatform_Click():                Unload Me: End Sub
Private Sub lblProductName_Click():             Unload Me: End Sub
Private Sub lblWarning_Click():                 Unload Me: End Sub

Private Sub lblCountDown_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub
Private Sub Timer1_Timer()
    lblCountDown.Caption = CountDown
    If CountDown < 0 Then Unload Me
    CountDown = CountDown - 1
End Sub
