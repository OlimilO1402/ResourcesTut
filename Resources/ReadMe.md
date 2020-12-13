* rc.exe und rcDll.dll lokalisieren oder holen aus dem Verzeichnis:
  - entweder Windows-SDK 
    C:\Programme\Windows Kits\10\bin\10.0.17763.0\x86\
    (sollte es einmal ein x86-VB geben dann aus dem X64-Verzeichnis)
    (C:\Programme\Windows Kits\10\bin\10.0.17763.0\x64\)
  - VB
    C:\Programme\Visual Studio\VB98\Wizards\

* die Ressourcenskriptdatei, MyRes.rc editieren und alle Änderungen 
  falls erforderlichen vornehmen.

* Konstanten sind in der Headerdatei MyRes.h definiert

* die Batchdatei MakeRes.bat doppelklicken.

Achtung Bug: 
Es wird kein Pfad akzeptiert, der mit einem T bzw. t beginnt.
oder Pfade immer mit Doppelbackslash ausrüsten, da /t irgend-
eine Bedeutung hat 

* seit manifest.exe.manifest
  bei Windows 7 folgende Codezeilen irgendwo in einem Modul im Projekt 
  hinzufügen:

  Public Declare Sub InitCommonControls Lib "comctl32.dll" () 
  'oder
  Private Declare Sub Application_EnableVisualStyles Lib "comctl32.dll" Alias "InitCommonControls" ()
  'irgendwo zum programmstart, im Startformular oder in Sub Main
  Private Sub Form_Initialize()
      Call InitCommonControls
  End Sub
