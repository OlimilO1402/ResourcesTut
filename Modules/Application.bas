Attribute VB_Name = "Application"
Option Explicit
Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Const MyExt   As String = ".ftext"
Public Const AppReg  As String = "RTCompany\ResourcesTutorial"
Public Const AppName As String = "ResourcesTutorial"
Public Const FIconId As Byte = 1 'for file-icon, prog-Icon-Id = 0
Public Const DefaultFileName As String = "RTDocument"

Public CurrentUserName As String

Sub Main()
    
    Settings.Load
    
    If Settings.ShowLoginAtStartup Then
        frmLogin.Show vbModal
    End If
    
    If Settings.ShowSplashAtStartup Then
        frmSplash.Show vbModeless, MDIMain
    End If
    
    If Settings.ShowTippsAtStartup Then
        frmTipp.Show vbModeless, MDIMain
    End If
    
    Dim r As ApiRect: r = Settings.MDIMainRect
    MDIMain.Move r.Left, r.Top, r.Width, r.Height
    
    MDIMain.Show
    MDIMain.WindowState = Settings.MDIMainWindowState
    
    If Len(Command) Then
        MDIMain.FileOpen Command
    Else
        MDIMain.FileNew
    End If

End Sub

Public Sub Terminate() 'called from MDIMain.Form_Unload
    'we also have to save all opened and changed files!
    'how to track data "changed"?
    'maybe the Undo&Redo could help, we must track file savings with UndoRedo
    'in some apps Undo will be cleared by file save, maybe we do not have to do this
    'simply by tracking a version-variable or a changedsincelastsaving-variable in Undo&Redo
    Settings.MDIMainWindowState = MDIMain.WindowState
    Settings.SetMDIMainRect MDIMain
    Settings.Save
End Sub

Public Function IsValidFileExt(ext As String) As Boolean
    IsValidFileExt = StrComp(ext, MyExt, vbTextCompare) = 0
    If IsValidFileExt Then Exit Function
    'hier evtl weitere Dateiformate prüfen falls true dann gleich raus
End Function

Public Sub RegisterExt()
    RegisterShellFileTypes MyExt, AppReg, AppName, App.path & "\" & AppName & ".exe", FIconId
End Sub

Public Sub UnRegisterExt()
    UnRegisterShellFileTypes MyExt, AppReg
End Sub

Public Function GetFilter() As String
    GetFilter = MyExt & "-Dateien [*" & MyExt & "]|*" & MyExt & "|Textdateien [*.txt]|*.txt|Alle Dateien [*.*]|*.*"
End Function

Public Function OpenFileName_ShowDlg(ByRef pfn_inout As String) As VbMsgBoxResult
Try: On Error GoTo Catch
    With MDIMain.FileDlg
        .InitDir = App.path
        If Len(pfn_inout) Then
            .FileName = pfn_inout
        End If
        .Filter = GetFilter
        .CancelError = True
        .ShowOpen
        pfn_inout = .FileName
    End With
    OpenFileName_ShowDlg = vbOK
    Exit Function
Catch: OpenFileName_ShowDlg = vbCancel
End Function

Public Function SaveFileName_ShowDlg(ByRef pfn_inout As String) As VbMsgBoxResult
Try: On Error GoTo Catch
    With MDIMain.FileDlg
        .InitDir = App.path
        If Len(pfn_inout) Then
            .FileName = pfn_inout
        End If
        .Filter = GetFilter
        .CancelError = True
        .ShowSave
        pfn_inout = .FileName
    End With
    
    'Settings.MRUFiles_Add pfn_inout
    
    SaveFileName_ShowDlg = vbOK
    Exit Function
Catch: SaveFileName_ShowDlg = vbCancel
End Function

