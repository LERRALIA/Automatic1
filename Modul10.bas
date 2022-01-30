Attribute VB_Name = "Modul10"
Option Explicit
Public Declare Function sevZIP_Init Lib "Zip32.dll" _
  Alias "InitZip" (ByVal sInit As String) As Boolean
  
Public Declare Sub sevZIP_SetLanguage Lib "Zip32.dll" _
  Alias "SetLanguage" (ByVal nLanguage As Long)

' Zip/Unzip-Funktionen
Public Declare Sub sevZIP_SetCompressionRate Lib "Zip32.dll" _
  Alias "SetCompressionRate" (ByVal nRate As Long)
  
Public Declare Function sevZIP_ZipFolderEx Lib "Zip32.dll" _
  Alias "ZipFolderEx" ( _
  ByVal sZipFile As String, _
  ByVal sSourcePath As String, _
  ByVal sFileSpec As String, _
  ByVal nSubFolder As Long, _
  ByVal sPassword As String, _
  ByVal nOverwrite As Long, _
  ByVal hStatus As Long) As Long

Public Declare Function sevZIP_ZipFile Lib "Zip32.dll" _
  Alias "ZipFile" ( _
  ByVal sZipFile As String, _
  ByVal sSourceFile As String, _
  ByVal sPassword As String, _
  ByVal nOverwrite As Long, _
  ByVal hStatus As Long) As Long
  
  '   Mögliche Rückgabewerte
' 0: Zip-File ist OK
' 1: Zip-File ist beschädigt oder enthält keine Dateien
' 2: Zip-File existiert nicht
' 3: Ungültiges Zip - File
Public Declare Function sevZIP_CheckZipFile Lib "Zip32.dll" _
  Alias "CheckZipFile" ( _
  ByVal sZipFile As String, _
  ByVal sPassword As String) As Long
  
Public Declare Function sevZIP_UnzipEx Lib "Zip32.dll" _
  Alias "UnZipEx" ( _
  ByVal sZipFile As String, _
  ByVal sDestPath As String, _
  ByVal sFileSpec As String, _
  ByVal nSubFolder As Long, _
  ByVal sPassword As String, _
  ByVal nOverwriteState As Long, _
  ByVal hStatus As Long) As Long

Public Sub Zip_Files(sPassword As String, sSourceFile As String, sZipFile As String, txtStatus As TextBox)
     On Error GoTo LOKAL_ERROR
  Dim nOverwrite As Long
  Dim nResult As Long

  ' Spracheinstellung und Kompressionsrate festlegen
  sevZIP_SetLanguage 1  ' Deutsch
  sevZIP_SetCompressionRate 5 'hoch ist 9 ' Val(Left$(cmbCompressType.Text, 1))
  
  ' Erste Datei, die gepackt werden soll
'  sSourceFile = lstFiles.list(0)
  
  ' Name und Ort des Zip-Archivs
'  sZipFile = txtZIPArchiv.Text
  
  ' Pfadangaben speichern?
'  If chkFolderLocation.Value <> 0 Then
'    sevZIP_SetRootDir ""
'  Else
'    sPath = Left$(sSourceFile, InStrRev(sSourceFile, "\") - 1)
'    sevZIP_SetRootDir sPath
'  End If

  ' ======================================================
  ' Für die Anzeige eines Fortschrittbalkens benötigen wir
  ' eine unsichtbare Textbox (txtStatus), sowie das
  ' ProgressBar-Control aus den Windows Commons Controls.
  '
  ' Wird beim Aufruf der Zip-Funktion ein gültiges
  ' Fenster-/ oder TextBox-Handle angegeben, so schreibt
  ' die DLL den aktuellen Fortschritt in dieses Fenster /
  ' TextBox
  '
  ' In unserem Beispiel verwenden wir eine unsichtbare
  ' TextBox, bei der die sevZIP32.DLL dann autom. das
  ' Change-Event auslöst, so dass wir den aktuellen
  ' Prozentsatz in der ProgressBar anzeigen können.
  ' ======================================================
  
  ' Passwort
'  sPassword = txtPassword.Text
  
  ' ZIP-Archiv überschreiben, falls bereits vorhanden
  nOverwrite = 1
  
  ' ZIP-Vorgang jetzt starten
  nResult = sevZIP_ZipFile(sZipFile, sSourceFile, _
    sPassword, nOverwrite, txtStatus.hwnd)
    
  ' Jetzt alle weiteren Dateien ins Archiv packen
'  With lstFiles
'    For i = 1 To .ListCount - 1
'      nResult = sevZIP_ZipAddFile(sZipFile, .list(i), _
'        2, txtStatus.hwnd)
'    Next i
'  End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "Zip_Files"
    Fehler.gsFehlertext = "Beim Zippen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Zip_Folder(sPassword As String, sFolder As String, sZipFile As String, txtStatus As TextBox)
On Error GoTo LOKAL_ERROR

    zipDllcheck

    Dim nOverwrite As Long
    Dim nSubFolder As Long
    Dim nResult As Long
    
    ' Spracheinstellung und Kompressionsrate festlegen
    sevZIP_SetLanguage 1  ' Deutsch
    sevZIP_SetCompressionRate 9 'hohe rate
    
    ' Ordner, dessen Dateien gepackt werden sollen
    '  sFolder = txtSource.Text
      
    ' Name und Ort des Zip-Archivs
    '  sZipFile = txtZIPArchiv.Text
      
    ' Pfadangaben speichern?
    '  If chkFolderLocation.Value <> 0 Then
    '    sevZIP_SetRootDir ""
    '  Else
    '    sevZIP_SetRootDir sFolder
    '  End If
      
    ' ======================================================
    ' Für die Anzeige eines Fortschrittbalkens benötigen wir
    ' eine unsichtbare Textbox (txtStatus), sowie das
    ' ProgressBar-Control aus den Windows Commons Controls.
    '
    ' Wird beim Aufruf der Zip-Funktion ein gültiges
    ' Fenster-/ oder TextBox-Handle angegeben, so schreibt
    ' die DLL den aktuellen Fortschritt in dieses Fenster /
    ' TextBox
    '
    ' In unserem Beispiel verwenden wir eine unsichtbare
    ' TextBox, bei der die sevZIP32.DLL dann autom. das
    ' Change-Event auslöst, so dass wir den aktuellen
    ' Prozentsatz in der ProgressBar anzeigen können.
    ' ======================================================
    
    ' Passwort
      
      
    ' ZIP-Archiv überschreiben, falls bereits vorhanden
    nOverwrite = 1
    
    ' Unterordner einbeziehen?
    nSubFolder = 0 'Abs(chkSubFolder.Value = 1)
    
    ' ZIP-Vorgang jetzt starten
    nResult = sevZIP_ZipFolderEx(sZipFile, sFolder, _
      "*.*", nSubFolder, sPassword, nOverwrite, txtStatus.hwnd)
        
    '  MsgBox CStr(nResult) & " Datei(en) gezippt."
      
      
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "Zip_Folder"
    Fehler.gsFehlertext = "Beim Zippen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ShowProgress(picprogress As Object, _
      ByVal value As Long, _
      ByVal Min As Long, _
      ByVal Max As Long, _
      Optional ByVal bShowProzent As Boolean = True)
      
      On Error GoTo LOKAL_ERROR
    
      
      Dim pWidth As Long
      Dim intProz As Integer
      Dim strProz As String
      
      ' Farben
      Const progBackColor = vbYellow
      Const progForeColor = vbBlue
      Const progForeColorHighlight = vbBlue
      
      ' Plausibilitätsprüfungen
      If value < Min Then value = Min
      If value > Max Then value = Max
      
      ' Prozentwert ausrechnen
      If Max > 0 Then
        intProz = Int(value / Max * 100 + 0.5)
      Else
        intProz = 100
      End If
        
      With picprogress
        ' Prüfen, ob AutoReadraw=True
        If .AutoRedraw = False Then .AutoRedraw = True
        
        ' Inhalt löschen
        picprogress.cLS
        
        If value > 0 Then
        
          ' Balkenbreite
          pWidth = .ScaleWidth / 100 * intProz
          
          ' Balken anzeigen
          picprogress.Line (0, 0)-(pWidth, .ScaleHeight), _
            progBackColor, BF
            
          ' Prozentanzeige
          If bShowProzent Then
            strProz = CStr(intProz) & " %"
            .CurrentX = (.ScaleWidth - .TextWidth(strProz)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strProz)) / 2
          
            ' Vordergrundfarbe
            If pWidth >= .CurrentX Then
              .ForeColor = progForeColorHighlight
            Else
              .ForeColor = progForeColor
            End If
          
            picprogress.Print strProz
          End If
        End If
      End With
      
      Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "ShowProgress"
    Fehler.gsFehlertext = "Beim Zippen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub zipDllcheck()
    On Error GoTo LOKAL_ERROR
    
    Dim lWert       As Long
    Dim cSysPfad    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim cPfad       As String
    
    cPfad = App.Path  'Anwendungspfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    cSysPfad = Space$(255)
    lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
    cSysPfad = Left(cSysPfad, lWert)

    If Not Modul6.FindFile(cSysPfad, "Zip32.dll") Then
        'kopiere
        cQuelle = cPfad & "\Zip32.dll"
        cZiel = cSysPfad & "\Zip32.dll"
        lRet = CopyFile(cQuelle, cZiel, lfail)
    End If
    
    Dim bresult As Boolean
    bresult = sevZIP_Init("6416-5529-7984-2605")
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul10"
        Fehler.gsFunktion = "zipdllcheck"
        Fehler.gsFehlertext = "Beim der Dateiüberprüfung sevZip32.dll ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Function DLLcheckZentStart(sPfad As String, sDatei As String, lFilesizeapp As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lDateiDatum As Long
    Dim lFileSize As Long
    
    DLLcheckZentStart = False
    
    
    
    If FindFile(sPfad, sDatei) = True Then
'        lDateiDatum = FileDateTime(sPfad & "\" & sDatei)
        lFileSize = fnFileSize(sPfad & "\" & sDatei)
        
'        MsgBox lFilesizeapp & lFileSize
        If lFilesizeapp <> lFileSize Then
        
        Else
            DLLcheckZentStart = True
        End If
        
    End If
    
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "DLLcheckZentStart"
    Fehler.gsFehlertext = "Bei der Dateiüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Public Sub mailocxcheck()
'    On Error GoTo LOKAL_ERROR
'
'    Dim lWert       As Long
'    Dim cSysPfad    As String
'    Dim cQuelle     As String
'    Dim cZiel       As String
'    Dim lfail       As Long
'    Dim lRet        As Long
'    Dim cPfad       As String
'    Dim lDateiDatum As Long
'
'    cPfad = App.Path  'Anwendungspfad
'    If Right(cPfad, 1) = "\" Then
'        cPfad = Left(cPfad, Len(cPfad) - 1)
'    End If
'
'    cSysPfad = Space$(255)
'    lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
'    cSysPfad = Left(cSysPfad, lWert)
'
'    If Not Modul6.FindFile(cSysPfad, "sevMail32.ocx") Then
'        'kopiere
'
'        cQuelle = cPfad
'        cQuelle = ShortPath(cQuelle)
'        cQuelle = cQuelle & "\Mail32.ocx"
'
'        cZiel = cSysPfad
'        cZiel = ShortPath(cZiel)
'        cZiel = cZiel & "\sevMail32.ocx"
'
'        lRet = CopyFile(cQuelle, cZiel, lfail)
'    Else
'        'neu trotzdem überbügeln
'        cQuelle = cPfad
'        cQuelle = ShortPath(cQuelle)
'        cQuelle = cQuelle & "\Mail32.ocx"
'
'        cZiel = cSysPfad
'        cZiel = ShortPath(cZiel)
'        cZiel = cZiel & "\sevMail32.ocx"
'
'        lRet = CopyFile(cQuelle, cZiel, lfail)
'
''        lDateiDatum = FileDateTime(cSysPfad & "\sevMail32.ocx")
''        If lDateiDatum < 39303 Then
''           MsgBox "Die Datei sevMail32.ocx ist veraltet. Bitte die Hotline anrufen! 0511955910", vbInformation, "Winkiss Hinweis:"
''        End If
'
'    End If
'
'    If Not Modul6.FindFile(cSysPfad, "sevMail32.dep") Then
'        'kopiere
'
'        cQuelle = cPfad
'        cQuelle = ShortPath(cQuelle)
'        cQuelle = cQuelle & "\Mail32.dep"
'
'        cZiel = cSysPfad
'        cZiel = ShortPath(cZiel)
'        cZiel = cZiel & "\sevMail32.dep"
'
'        lRet = CopyFile(cQuelle, cZiel, lfail)
'
'    End If
'
'    If Not Modul6.FindFile(cSysPfad, "sevMail32.oca") Then
'        'kopiere
'
'        cQuelle = cPfad
'        cQuelle = ShortPath(cQuelle)
'        cQuelle = cQuelle & "\Mail32.oca"
'
'        cZiel = cSysPfad
'        cZiel = ShortPath(cZiel)
'        cZiel = cZiel & "\sevMail32.oca"
'
'        lRet = CopyFile(cQuelle, cZiel, lfail)
'
'    End If
'
'    doIt cSysPfad & "\sevMail32.ocx", True  'False für unregister
'
'    Exit Sub
'LOKAL_ERROR:
'    If err.Number = 53 Then
'        Exit Sub
'    Else
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul10"
'        Fehler.gsFunktion = "mailocxcheck"
'        Fehler.gsFehlertext = "Beim der Dateiüberprüfung sevMail32.ocx ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
'    End If
End Sub
Public Sub mailDLLcheck()
    On Error GoTo LOKAL_ERROR
    
    Dim lWert       As Long
    Dim cSysPfad    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim cPfad       As String
    Dim cAppPfad    As String
    Dim lDateiDatum As Long
    
    cPfad = App.Path  'Anwendungspfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    cAppPfad = App.Path  'Anwendungspfad
    If Right(cAppPfad, 1) = "\" Then
        cAppPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    cAppPfad = ShortPath(cAppPfad)
    
    cSysPfad = Space$(255)
    lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
    cSysPfad = Left(cSysPfad, lWert)

    If Not Modul6.FindFile(cSysPfad, "EASendMailObj.dll") Then
        'kopiere
        
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "\EASendMailObj.dll"
        
        cZiel = cSysPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "\EASendMailObj.dll"
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
    Else
        'neu trotzdem überbügeln
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "\EASendMailObj.dll"
        
        cZiel = cSysPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "\EASendMailObj.dll"
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
        
'        lDateiDatum = FileDateTime(cSysPfad & "\sevMail32.ocx")
'        If lDateiDatum < 39303 Then
'           MsgBox "Die Datei sevMail32.ocx ist veraltet. Bitte die Hotline anrufen! 0511955910", vbInformation, "Winkiss Hinweis:"
'        End If

    End If
    
    
    doIt cAppPfad & "\EASendMailObj.dll", True  'False für unregister
    doIt cSysPfad & "\EASendMailObj.dll", True  'False für unregister
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul10"
        Fehler.gsFunktion = "mailDLLcheck"
        Fehler.gsFehlertext = "Beim der Dateiüberprüfung sevMail32.ocx ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub zip32check()
    On Error GoTo LOKAL_ERROR
    
    Dim lWert       As Long
    Dim cSysPfad    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim cPfad       As String
    Dim lDateiDatum As Long
    
    cPfad = App.Path  'Anwendungspfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    cSysPfad = Space$(255)
    lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
    cSysPfad = Left(cSysPfad, lWert)

    If Not Modul6.FindFile(cSysPfad, "zip32.dll") Then
        'kopiere
        
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "\zip32.dll"
        
        cZiel = cSysPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "\zip32.dll"
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
    Else
        'neu trotzdem überbügeln
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "\zip32.dll"
        
        cZiel = cSysPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "\zip32.dll"
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
    End If
    
    doIt cSysPfad & "\zip32.dll", True  'False für unregister
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul10"
        Fehler.gsFunktion = "zip32check"
        Fehler.gsFehlertext = "Beim der Dateiüberprüfung sevMail32.ocx ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub systemdatcheck(sdatname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lWert       As Long
    Dim cSysPfad    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim cPfad       As String
    
    cPfad = gcPfad  'Anwendungspfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    If Modul6.FindFile(cPfad, sdatname) Then
        cSysPfad = Space$(255)
        lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
        cSysPfad = Left(cSysPfad, lWert)
        
        'kopiere
        cQuelle = cPfad & "\"
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & sdatname
        
        
        cZiel = cSysPfad & "\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & sdatname
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
            
    End If
    
    doIt cSysPfad & "\" & sdatname, True  'False für unregister
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "systemdatcheck"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Sub Gemeindatcheck(sdatname As String, breg As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lWert       As Long
    Dim cSysPfad    As String
    Dim cSysPfadD    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim cPfad       As String
    
    cPfad = gcPfad  'Anwendungspfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    If Modul6.FindFile(cPfad, sdatname) Then
        cSysPfad = "C:\Programme\Gemeinsame Dateien\Microsoft Shared\DAO"
        
        'kopiere
        cQuelle = cPfad & "\" & sdatname
        cZiel = cSysPfad & "\" & sdatname
        lRet = CopyFile(cQuelle, cZiel, lfail)
        
        cSysPfadD = "D:\Programme\Gemeinsame Dateien\Microsoft Shared\DAO"
        
        'kopiere
        cQuelle = cPfad & "\" & sdatname
        cZiel = cSysPfadD & "\" & sdatname
        lRet = CopyFile(cQuelle, cZiel, lfail)
        
       
            
    End If
    
    doIt cSysPfad & "\" & sdatname, breg  'False für unregister
    doIt cSysPfadD & "\" & sdatname, breg  'False für unregister
    
    
    
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "Gemeindatcheck"
    Fehler.gsFehlertext = "Bei der Dateiüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Sub Zip_Unzip(sPassword As String, sDestPath As String, sZipFile As String, txtStatus As TextBox)
    On Error GoTo LOKAL_ERROR
    
    zipDllcheck
    
'    Dim sZipFile As String
'    Dim sDestPath As String
'    Dim sPassword As String
    Dim nOverwriteState As Long
    Dim nResult As Long
    Dim nSubFolder As Long
    
    ' Spracheinstellung und Kompressionsrate festlegen
    sevZIP_SetLanguage 1  ' Deutsch
    
    ' Name und Ort des Zip-Archivs
'    sZipFile = txtZIPArchiv.Text
    
    ' Zielverzeichnis, in das die Datei(en) entpackt
    ' werden sollen
'    sDestPath = txtFolder.Text
    
    ' ======================================================
    ' Für die Anzeige eines Fortschrittbalkens benötigen wir
    ' eine unsichtbare Textbox (txtStatus), sowie das
    ' ProgressBar-Control aus den Windows Commons Controls.
    '
    ' Wird beim Aufruf der UnZip-Funktion ein gültiges
    ' Fenster-/ oder TextBox-Handle angegeben, so schreibt
    ' die DLL den aktuellen Fortschritt in dieses Fenster /
    ' TextBox
    '
    ' In unserem Beispiel verwenden wir eine unsichtbare
    ' TextBox, bei der die sevZIP32.DLL dann autom. das
    ' Change-Event auslöst, so dass wir den aktuellen
    ' Prozentsatz in der ProgressBar anzeigen können.
    ' ======================================================
    
    ' Passwort
'    sPassword = txtPassword.Text
    
    ' Überschreib-Status
'    If optOverwrite(0).Value Then
'      ' niemals überschreiben
'      nOverwriteState = 3
'    ElseIf optOverwrite(1).Value Then
'      ' immer überschreiben
      nOverwriteState = 2
'    Else
'      ' ggf. Hinweis anzeigen
'      nOverwriteState = 0
'    End If
    
    ' Verzeichnisstruktur übernehmen?
    nSubFolder = 0 'Abs(chkSubFolder.Value = 1)
    
    ' UnZip-Vorgang jetzt starten
    
    
    
    If sevZIP_CheckZipFile(sZipFile, "") > 0 Then
    
        If gbNacht = False Then
            MsgBox "Achtung!" & vbCrLf & "Ungültige oder beschädtigte Zip-Datei!" & vbCrLf & sZipFile
        End If
        Kill sZipFile
      
    Else
        nResult = sevZIP_UnzipEx(sZipFile, sDestPath, _
        "*.*", nSubFolder, sPassword, nOverwriteState, txtStatus.hwnd)
    End If

''    nResult = sevZIP_UnzipEx(sZipFile, sDestPath, _
''      "*.*", nSubFolder, sPassword, nOverwriteState, txtStatus.hwnd)
      
      
      
      
'    MsgBox CStr(nResult) & " Datei(en) entpackt."
    
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul10"
        Fehler.gsFunktion = "Zip_Unzip"
        Fehler.gsFehlertext = "Beim Zippen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Function Gibt_es_Termine_in_Zunkunft(cKdnr As String, Optional lBestimmtesDatum As Long) As Long
On Error GoTo LOKAL_ERROR

    Gibt_es_Termine_in_Zunkunft = 0
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    If cKdnr = "" Then Exit Function

    sSQL = " Select * from Termine where kundnr = " & cKdnr
    
    If lBestimmtesDatum > 0 Then
        sSQL = sSQL & " and datum = " & lBestimmtesDatum
    Else
        sSQL = sSQL & " and datum >= " & CLng(DateValue(Now))
    End If
    

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Gibt_es_Termine_in_Zunkunft = 1 'rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul10"
    Fehler.gsFunktion = "Gibt_es_Termine_in_Zunkunft"
    Fehler.gsFehlertext = "Im Programmteil Kunde suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function



