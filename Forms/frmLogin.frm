VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Login"
   ClientHeight    =   1455
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1340
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1340
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2445
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Username:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnCancel_Click()
    'Application.CurrentUserName = txtUserName.Text
    Application.CurrentUserName = "Oliver"
    Unload Me
End Sub

Private Sub BtnOK_Click()
    'check password
    Dim pw As String: pw = txtPassword.Text
    If pw = "secret" Or pw = "geheim" Then
        Application.CurrentUserName = txtUserName.Text
        Unload Me
    Else
        MsgBox "Tipp: the password is secret!", , "Login"
        txtPassword.SetFocus
        'SendKeys "{Home}+{End}"
    End If
End Sub
