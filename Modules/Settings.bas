Attribute VB_Name = "Settings"
Option Explicit
Public Type ApiRect
    Left   As Long
    Top    As Long
    Width  As Long
    Height As Long
End Type
Public Enum ELanguage
    Langg_None
    Langg_English
    Langg_German
    Langg_Hrvatsk
    Langg_Greek
End Enum
Public ShowSplashAtStartup  As Boolean
Public ShowTippsAtStartup   As Boolean
Public ShowLoginAtStartup   As Boolean
Public FileIconIsRegistered As Boolean
Public Language             As ELanguage
Public MDIMainWindowState   As FormWindowStateConstants
Public MDIChildWindowState  As FormWindowStateConstants
Public MDIMainRect          As ApiRect
Public SettingsAreDeleted   As Boolean
Private Const Section As String = "Settings"

Private Function ApiRect_ToStr(r As ApiRect) As String
    ApiRect_ToStr = r.Left & "," & r.Top & "," & r.Width & "," & r.Height
End Function
Private Function ApiRect_Parse(s As String) As ApiRect
    Dim sa() As String: sa = Split(s, ",")
    With ApiRect_Parse
        .Left = sa(0): .Top = sa(1): .Width = sa(2): .Height = sa(3)
    End With
End Function

Public Sub SetMDIMainRect(F As MDIMain)
    With MDIMainRect
        .Left = F.Left
        .Top = F.Top
        .Width = F.Width
        .Height = F.Height
    End With
End Sub

'Public Function ELanguage_ToStr(el As ELanguage) As String
'    Dim s As String
'    Select Case el
'    Case ELanguage.Langg_English: s = "English"
'    Case ELanguage.Langg_German:  s = "German"
'    Case ELanguage.Langg_Hrvatsk: s = "Hrvatsk"
'    Case ELanguage.Langg_Greek:   s = "Greek"
'    End Select
'    ELanguage_ToStr = s
'End Function

Public Sub Load()
    
    Dim apnam As String: apnam = Application.AppName
    Dim scnam As String: scnam = "Settings"
    ShowSplashAtStartup = GetSetting(apnam, scnam, "ShowSplashAtStartup", True)
    ShowTippsAtStartup = GetSetting(apnam, scnam, "ShowTippsAtStartup", True)
    ShowLoginAtStartup = GetSetting(apnam, scnam, "ShowLoginAtStartup", True)
    
    FileIconIsRegistered = GetSetting(apnam, scnam, "FileIconIsRegistered", False)
    MDIMainWindowState = GetSetting(apnam, scnam, "MDIMainWindowState", FormWindowStateConstants.vbNormal)
    MDIMainRect = ApiRect_Parse(GetSetting(apnam, scnam, "MDIMainWindowRect", "105,105,11505,7815")) 'L: 105; T: 105; W: 11505; H: 7815
    MDIChildWindowState = GetSetting(apnam, scnam, "MDIChildWindowState", FormWindowStateConstants.vbMaximized)
    Language = GetSetting(apnam, scnam, "Language", ELanguage.Langg_English)
    
End Sub

Public Sub Save()
    
    If SettingsAreDeleted Then Exit Sub
    Dim apnam As String: apnam = Application.AppName
    SaveSetting apnam, Section, "ShowSplashAtStartup", ShowSplashAtStartup
    SaveSetting apnam, Section, "ShowTippsAtStartup", ShowTippsAtStartup
    SaveSetting apnam, Section, "ShowLoginAtStartup", ShowLoginAtStartup
    SaveSetting apnam, Section, "FileIconIsRegistered", FileIconIsRegistered
    SaveSetting apnam, Section, "MDIMainWindowState", MDIMainWindowState
    SaveSetting apnam, Section, "MDIMainWindowRect", ApiRect_ToStr(MDIMainRect)
    SaveSetting apnam, Section, "MDIChildWindowState", MDIChildWindowState
    SaveSetting apnam, Section, "Language", Language
    
End Sub

Public Sub Delete()

    'all settings stored in the registry in key:
    'HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\<AppName>\<Section>\<Key>
    'will be deleted
    
    Dim apnam As String: apnam = Application.AppName
    Dim path  As String: path = "SOFTWARE\VB and VBA Program Settings\" & apnam
    
    Registry.Init
    
    Registry.RootKey = HKEY_CURRENT_USER
    Registry.DeleteKey path & "\" & Section
    Registry.CloseKey
    
    Registry.RootKey = HKEY_CURRENT_USER:
    Registry.DeleteKey path
    Registry.CloseKey
    
    SettingsAreDeleted = True
    
End Sub
