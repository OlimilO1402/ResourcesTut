VERSION 5.00
Begin VB.Form frmTipp 
   Caption         =   "Tips und Tricks"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTippsAtStartup 
      Caption         =   "&Tipps beim Starten anzeigen"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2415
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&N�chster Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTipp.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wu�ten Sie schon.."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTipp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Tip-Datenbank im Speicher.
Dim Tips As New Collection

' Name der Tip-Datei
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in der Tip-Auflistung, die momentan angezeigt wird.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Einen Tip willk�rlich ausw�hlen.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Oder die Tips der Reihenfolge nach durchgehen.

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Tip anzeigen.
    DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Jeder Tip wird aus der Datei eingelesen.
    Dim InFile As Integer   ' Descriptor f�r Datei.
    
    ' N�chsten freien Datei-Descriptor abrufen.
    InFile = FreeFile
    
    ' Sicherstellen, da� eine Datei angegeben wurde.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Sicherstellen, da� die Datei vorhanden ist, bevor sie ge�ffnet wird.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Auflistung aus einer Text-Datei lesen.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Tips willk�rlich anzeigen.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Dim ShowAtStartup As Long
    
    ' Feststellen, ob das Dialogfeld beim Start angezeigt werden soll
    'ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    'If ShowAtStartup = 0 Then
    '    Unload Me
    '    Exit Sub
    'End If
        
    ' Kontrollk�stchen festlegen. Hierdurch wird der Wert in die Registrierung geschrieben
    Me.chkLoadTippsAtStartup.Value = Abs(Settings.ShowTippsAtStartup)
    
    ' Randomisieren beginnen
    Randomize
    
    ' Tip-Datei lesen und einen Tip willk�rlich anzeigen.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "Die Datei " & TIP_FILE & " wurde nicht gefunden? " & vbCrLf & vbCrLf & _
           "Textdatei mit dem Namen " & TIP_FILE & " unter Verwendung von NotePad mit 1 Tip pro Zeile erstellen. " & _
           "Dann im selben Verzeichnis wie die Anwendung ablegen. "
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Settings.ShowTippsAtStartup = chkLoadTippsAtStartup.Value = vbChecked
End Sub
