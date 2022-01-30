Attribute VB_Name = "Modul8"
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const ANYSIZE_ARRAY = 1
Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8
Public Const SE_PRIVILEGE_ENABLED As Long = 2
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8
Public Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Private Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"

Public Type LARGE_INTEGER
   lowpart As Long
   highpart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
   pLuid As LARGE_INTEGER
   Attributes As Long
End Type

Public Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Type OSVERSIONINFO ' für den Aufruf des GetVersionEx-API
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInternetHandle As Long) As Boolean
Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal lpszCallerName As String, _
ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As _
String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrlA Lib "wininet.dll" (ByVal hOpen As Long, ByVal _
sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, _
ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal _
sBuffer As String, ByVal lNumBytesToRead As Long, Bytes As Long) As Integer
Public Enum InternetOpenType
  IOTPreconfig = 0
  IOTDirect = 1
  IOTProxy = 3
End Enum
Public Function OpenURL(ByVal URL As String, Optional ByVal OpenType As _
InternetOpenType = IOTPreconfig) As String
  Const INET_RELOAD = &H80000000
  Dim hInet As Long
  Dim hURL As Long
  Dim Buffer As String * 2048
  Dim Bytes As Long
  'Inet-Connection öffnen:
  hInet = InternetOpenA("VB-Tec:INET", OpenType, vbNullString, vbNullString, 0)
  hURL = InternetOpenUrlA(hInet, URL, vbNullString, 0, INET_RELOAD, 0)
  'Daten sammeln:
  Do
    InternetReadFile hURL, Buffer, Len(Buffer), Bytes
    If Bytes = 0 Then Exit Do
    OpenURL = OpenURL & Left$(Buffer, Bytes)
  Loop
  'Inet-Connection schließen:
  InternetCloseHandle hURL
  InternetCloseHandle hInet
End Function

'Private Declare Function ShellExecute Lib "shell32.dll" _
'  Alias "ShellExecuteA" ( _
'  ByVal hwnd As Long, _
'  ByVal lpOperation As String, _
'  ByVal lpFile As String, _
'  ByVal lpParameters As String, _
'  ByVal lpDirectory As String, _
'  ByVal nShowCmd As Long) As Long
 
' Diese nachfolgende Prozedur aktiviert den im
' System registrierten Standard-Browser und lädt
' die durch URL angegebene Internetadresse
Public Sub URLGoTo(ByVal hwnd As Long, ByVal URL As String)
 
  ' hWnd: Das Fensterhandle des
  ' aufrufenden Formulars
 
  Screen.MousePointer = 11
  
'  vbNullString
  
'  Call ShellExecute(hwnd, "Open", URL, "", "", SW_MINIMIZE)
  
  Call ShellExecute(hwnd, "Open", URL, "", "", 1)
  Screen.MousePointer = 0
End Sub
Public Function ermMENGE44(sArt As String, cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermMENGE44 = 0#

sSQL = "Select sum(Menge) as maxi from Kassjour where "
sSQL = sSQL & " adate between  " & cVon & " And " & cBis
sSQL = sSQL & " and artnr = " & sArt

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermMENGE44 = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermMENGE44"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermKundenName(sKUNDNR As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermKundenName = ""

sSQL = "Select name from kunden where "
sSQL = sSQL & "  kundnr = " & sKUNDNR

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!name) Then
        ermKundenName = rsrs!name
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermKundenName"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermBedKuerzel(sKUNDNR As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermBedKuerzel = "0"

sSQL = "Select name from kunden where "
sSQL = sSQL & "  kundnr = " & sKUNDNR

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!name) Then
        ermBedKuerzel = UCase(Left(rsrs!name, 5))
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermBedKuerzel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermBedKundnr(sbed As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermBedKundnr = "0"

sSQL = "Select kundnr from bedterm where "
sSQL = sSQL & "  bednu = " & sbed

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!Kundnr) Then
        ermBedKundnr = rsrs!Kundnr
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermBedKundnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Public Function ermKWERT(sArt As String, cWert As String) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermKWERT = 0#

sSQL = "Select " & cWert & " as maxi from Artikel where "
sSQL = sSQL & "  artnr = " & sArt

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermKWERT = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermKWERT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermGutschVkdat(sGutsch As String) As String
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermGutschVkdat = ""

sSQL = "Select dat_ausg from Gutsch where "
sSQL = sSQL & "  gutschnr = " & sGutsch

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!DAT_AUSG) Then
        ermGutschVkdat = rsrs!DAT_AUSG
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermGutschVkdat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Public Function ermGutschkkart(sGutsch As String, sVKdat As String) As String
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset
Dim cSuch As String

ermGutschkkart = ""

cSuch = "GUTSCHEIN " & Trim(sGutsch)

sSQL = "Select kk_art from kassjour where Bezeich like '" & cSuch & "*' "
sSQL = sSQL & " and adate = " & CLng(DateValue(sVKdat))
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!kk_art) Then
        ermGutschkkart = rsrs!kk_art
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermGutschkkart"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub BringFarbeInsSpielforKunden(sTab As String, db As Database)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lFarbert    As Long
    Dim lFarbert2   As Long
    Dim i           As Integer

    sSQL = "Update " & sTab & " set farbnr = 0 "
    sSQL = sSQL & "  where farbnr is null "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & sTab & " inner join FarbKu on " & sTab & ".farbnr = FarbKu.farbnr"
    sSQL = sSQL & " set " & sTab & ".farbtext = FarbKu.farbtext "
    db.Execute sSQL, dbFailOnError
     
    sSQL = "Update " & sTab & " set farbtext = 'ohne Beschreibung' "
    sSQL = sSQL & "  where farbtext is null "
    db.Execute sSQL, dbFailOnError
     
    sSQL = "Update " & sTab & " set farbtext = 'ohne Kennzeichen' "
    sSQL = sSQL & "  where farbnr = 0 "
    db.Execute sSQL, dbFailOnError
     
    For i = 1 To 9
        lFarbert = CDec(glfarbe(i))
        lFarbert2 = vbBlack
        sSQL = "Update " & sTab & " set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i
        db.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 9
        lFarbert = CDec(glfarbe2(i))
        lFarbert2 = vbBlack
        sSQL = "Update " & sTab & " set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i + 10
        db.Execute sSQL, dbFailOnError
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "BringFarbeInsSpielforKunden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub BringFarbeInsSpiel(sTab As String, db As Database)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lFarbert    As Long
    Dim lFarbert2   As Long
    Dim i           As Integer

    sSQL = "Update " & sTab & " set farbnr = 0 "
    sSQL = sSQL & "  where farbnr is null "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & sTab & " set farbtext = 'neue Artikel' "
    sSQL = sSQL & "  where farbnr = 98 "
    db.Execute sSQL, dbFailOnError
     
    sSQL = "Update " & sTab & " set farbtext = 'nicht geliefert' "
    sSQL = sSQL & "  where farbnr = 95 "
    db.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update " & sTab & " inner join Farbmerk on " & sTab & ".farbnr = farbmerk.farbnr"
    sSQL = sSQL & " set " & sTab & ".farbtext = farbmerk.farbtext "
    db.Execute sSQL, dbFailOnError
     
    sSQL = "Update " & sTab & " set farbtext = 'ohne Beschreibung' "
    sSQL = sSQL & "  where farbtext is null "
    db.Execute sSQL, dbFailOnError
     
    sSQL = "Update " & sTab & " set farbtext = 'ohne Kennzeichen' "
    sSQL = sSQL & "  where farbnr = 0 "
    db.Execute sSQL, dbFailOnError
     
    For i = 1 To 9
        lFarbert = CDec(glfarbe(i))
        lFarbert2 = vbBlack
        sSQL = "Update " & sTab & " set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i
        db.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 9
        lFarbert = CDec(glfarbe2(i))
        lFarbert2 = vbBlack
        sSQL = "Update " & sTab & " set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i + 10
        db.Execute sSQL, dbFailOnError
    Next i
    
    lFarbert = vbRed
    lFarbert2 = vbWhite
    
    sSQL = "Update " & sTab & " set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  98 "
    db.Execute sSQL, dbFailOnError
     
    lFarbert = vbBlack
    lFarbert2 = vbBlue
    
    sSQL = "Update " & sTab & " set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  95 "
    db.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "BringFarbeInsSpiel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub LoescheArtikelSofort(cArtNr As String, sDelText As String)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim rsrs As Recordset
    
    SicherInArtikelsic CLng(cArtNr)
    
    cSQL = "Select * from ARTIKEL where ARTNR = " & cArtNr & " "
    cSQL = cSQL & " and ( aufdat < " & CLng(DateValue(Now)) - 56
    cSQL = cSQL & " or aufdat is null )"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
    Else
        rsrs.Close: Set rsrs = Nothing
        Exit Sub
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    schreibeProtokollgArtikel "Artikel: " & cArtNr & " " & ErmittleDetails(cArtNr) & " " & sDelText

    cSQL = "Delete from ARTIKEL where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Artlief where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "LoescheArtikelSofort"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub fülleSpalte(cbox As ComboBox, sSpalte As String, sTab As String, sOrder As String, sAnzeigetext As String, cFormatart As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim dateDat As Date
    
    If Not NewTableSuchenDBKombi(sTab, gdBase) Then
        Exit Sub
    End If
    
    cbox.Clear
    If sAnzeigetext <> "" Then
        cbox.AddItem sAnzeigetext
    End If
    
    cbox.Text = sAnzeigetext
    
    sSQL = "Select distinct(" & sSpalte & ")as maxi from " & sTab & " order by " & sOrder
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!maxi) Then
                cFeld = rsrs!maxi
                Select Case cFormatart
                Case "d"
                    cFeld = Format(cFeld, "###0.00")
                Case "D"
                    dateDat = rsrs!maxi
                    cFeld = DateValue(dateDat)
                Case Else
                    cFeld = cFeld
                End Select
                cbox.AddItem cFeld
            End If
            cFeld = ""
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "fülleSpalte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub fülleSpalte_KL(cbox As ComboBox, sSpalte As String, sTab As String, sOrder As String, sAnzeigetext As String, cFormatart As String, cwhere As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim dateDat As Date
    
    
    
    
    
    
    
    
    
    
    Dim stConnect As String
    
    If gsKL_DSN <> "" Then
        stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    Else
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    End If
    
    Dim dbEAN As DAO.Database
    Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
    
    
    If Not NewTableSuchenDBKombi(sTab, dbEAN) Then
        Exit Sub
    End If
    
    cbox.Clear
    If sAnzeigetext <> "" Then
        cbox.AddItem sAnzeigetext
    End If
    
    cbox.Text = sAnzeigetext
    
    sSQL = "Select distinct(" & sSpalte & ")as maxi from " & sTab & " " & cwhere
    sSQL = sSQL & " order by " & sOrder
    Set rsrs = dbEAN.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!maxi) Then
                cFeld = rsrs!maxi
                Select Case cFormatart
                Case "d"
                    cFeld = Format(cFeld, "###0.00")
                Case "D"
                    dateDat = rsrs!maxi
                    cFeld = DateValue(dateDat)
                Case Else
                    cFeld = cFeld
                End Select
                cbox.AddItem cFeld
            End If
            cFeld = ""
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dbEAN.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "fülleSpalte_KL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub fülledistinctTabelle(sSpalte As String, sSpalte2 As String, sIntoTab As String, sTab As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    loeschNEW sIntoTab, gdBase
    
    If sSpalte2 <> "" Then
        sSQL = "Select distinct(" & sSpalte & ")as " & sSpalte & "d, " & sSpalte2 & " into " & sIntoTab & " from " & sTab
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Select distinct(" & sSpalte & ")as " & sSpalte & "d into " & sIntoTab & " from " & sTab
        gdBase.Execute sSQL, dbFailOnError
    End If

    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul8"
        Fehler.gsFunktion = "fülledistinctTabelle"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub fülleSpaltemitKrit(cbox As ComboBox, sSpalte As String, sTab As String, sOrder As String, sAnzeigetext As String, cFormatart As String, cKrit As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs45 As Recordset
    Dim cFeld As String
    Dim dateDat As Date
    Dim lcount As Long
    
    cbox.Clear
    If sAnzeigetext <> "" Then
        cbox.AddItem sAnzeigetext
    End If
    
    cbox.Text = sAnzeigetext
    
    sSQL = "Select distinct(" & sSpalte & ") as maxi  from " & sTab & " " & cKrit & " order by " & sOrder
    Set rsrs45 = gdBase.OpenRecordset(sSQL)
    If Not rsrs45.EOF Then
        rsrs45.MoveLast
        lcount = rsrs45.RecordCount
        rsrs45.MoveFirst
        Do While Not rsrs45.EOF
            If Not IsNull(rsrs45!maxi) Then
                cFeld = rsrs45!maxi
                Select Case cFormatart
                Case "d"
                    cFeld = Format(cFeld, "###0.00")
                Case "D"
                    dateDat = rsrs45!maxi
                    cFeld = DateValue(dateDat)
                Case Else
                    cFeld = cFeld
                End Select
                cbox.AddItem cFeld
                If lcount = 1 Then
                    cbox.Text = cFeld
                End If
            End If
            cFeld = ""
        rsrs45.MoveNext
        Loop
    End If
    rsrs45.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "fülleSpaltemitKrit"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Sub füllefil(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String
    Dim cFeld As String
    
    cbox.Clear
    cbox.AddItem "bitte wählen"
    cbox.Text = "bitte wählen"
    
    sSQL = "Select * from filialen order by filialnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
            
                cFeld = rsrs!FILIALNR
                cSatz = cSatz & Space(3 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!FILIALNAME) Then
                    cFeld = rsrs!FILIALNAME
                    cSatz = cSatz & Space(2) & cFeld
                    cbox.AddItem cSatz
                End If
            End If
            cSatz = ""
            cFeld = ""
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "füllefil"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub filcboBediener(cboBed As ComboBox)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sTemp As String
    Dim cSatz As String
    Dim cFeld As String
    Dim counter As Long
    Dim lAnzahl As Long
    
    sSQL = "Select bednu,bedname from bedname order by bedname"
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboBed.Clear
    cboBed.AddItem "alle Bediener"
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!BEDNU) Then
            
                cFeld = rs!BEDNU
                cSatz = cSatz & Space(4 - Len(cFeld)) & cFeld
                
                If Not IsNull(rs!bedname) Then
                    cFeld = rs!bedname
                    cSatz = cSatz & Space(2) & cFeld
                    cboBed.AddItem cSatz
                    cSatz = ""
                    cFeld = ""
                End If
            End If
        rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    cboBed.Text = "alle Bediener"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "filcboBediener"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ErmittleDetails(art As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    ErmittleDetails = ""
    
    cSQL = "Select * from ARTIKEL where ARTNR = " & art & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        
        If Not IsNull(rsrs!BEZEICH) Then
            ErmittleDetails = rsrs!BEZEICH
        End If
        
        If Not IsNull(rsrs!linr) Then
            ErmittleDetails = ErmittleDetails & vbCrLf & "Lieferant: " & rsrs!linr
        End If
    
        If Not IsNull(rsrs!LPZ) Then
            ErmittleDetails = ErmittleDetails & " Linie: " & rsrs!LPZ
        End If
    
        If Not IsNull(rsrs!LIBESNR) Then
            ErmittleDetails = ErmittleDetails & " BestellNr: " & rsrs!LIBESNR
        End If
    
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ErmittleDetails"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SetTime(Datum As Date) As Boolean
    On Error GoTo LOKAL_ERROR
    
  Dim ret As Long
  Dim hToken As Long
  Dim tkp As TOKEN_PRIVILEGES
  Dim tkpOld As TOKEN_PRIVILEGES
  Dim fOk As Boolean
  Dim Time As SYSTEMTIME

' Überprüfen, ob Window NT ausgeführt wird.
   If IsWindowsNT() Then
' Windows NT wird ausgeführt.
' Sicherheitsüberprüfungen und
' Veränderungen sind jetzt notwendig,
' um sicherzustellen, daß das Token
' vorhanden ist, das einen Neustart zuläßt.
      If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
         ret = LookupPrivilegeValue(vbNullString, SE_SYSTEMTIME_NAME, tkp.Privileges(0).pLuid)
         tkp.PrivilegeCount = 1
         tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
         fOk = AdjustTokenPrivileges(hToken, 0, tkp, LenB(tkpOld), tkpOld, ret)
      End If
   Else
' Win95/98 wird ausgeführt. Keine Aktion ist notwendig.
      fOk = True
   End If
   If fOk Then
      Time.wSecond = Val(Format(Datum, "ss"))
      Time.wMinute = Val(Mid(Format(Datum, "long time"), InStr(Format(Datum, "long time"), ":") + 1))
      Time.wHour = Val(Format(Datum, "hh")) - 1
      Time.wDay = Val(Format(Datum, "d"))
      Time.wMonth = Val(Format(Datum, "m"))
      Time.wYear = Val(Format(Datum, "yyyy"))
      
      SetTime = (SetSystemTime(Time) <> 0)
   End If
   
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "SetTime"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function

Public Function SystemDown() As Boolean
On Error GoTo LOKAL_ERROR

  Dim ret As Long
  Dim hToken As Long
  Dim tkp As TOKEN_PRIVILEGES
  Dim tkpOld As TOKEN_PRIVILEGES
  Dim fOKRunterfahren As Boolean
  Const sSHUTDOWN As String = SE_SHUTDOWN_NAME

' Überprüfen, ob Windows NT ausgeführt wird.
   If IsWindowsNT() Then
' Windows NT wird ausgeführt.
' Sicherheitsüberprüfungen und
' Veränderungen sind jetzt notwendig,
' um sicherzustellen, daß das Token
' vorhanden ist, das einen Neustart zuläßt.
      If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
         ret = LookupPrivilegeValue(vbNullString, sSHUTDOWN, tkp.Privileges(0).pLuid)
         tkp.PrivilegeCount = 1
         tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
         fOKRunterfahren = AdjustTokenPrivileges(hToken, 0, tkp, LenB(tkpOld), tkpOld, ret)
      End If
   Else
' Win95/98 wird ausgeführt. Keine Aktion ist notwendig.
      fOKRunterfahren = True
   End If
   If fOKRunterfahren Then
      'SystemDown = (ExitWindowsEx(EWX_SHUTDOWN, 0) <> 0)
      If IsWindowsNT() Then
      SystemDown = (ExitWindowsEx(EWX_POWEROFF, 0) <> 0)
    Else
      SystemDown = (ExitWindowsEx(EWX_SHUTDOWN, 0) <> 0)
    End If
   
   End If
   
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "SystemDown"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

'-----------------------------------------------------------
' FUNKTION: IsWindowsNT
'
' Liefert "True", falls dieses Programm unter
' Windows NT ausgeführt wird.
'-----------------------------------------------------------
'
Public Function IsWindowsNT() As Boolean
On Error GoTo LOKAL_ERROR

   Const dwMaskNT = &H2&
   IsWindowsNT = (GetWinPlatform() And dwMaskNT)
   
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "IsWindowsNT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

'----------------------------------------------------------
' FUNKTION: GetWinPlatform
' Aktuelle Windows-Plattform ermitteln.
' ---------------------------------------------------------
Public Function GetWinPlatform() As Long
On Error GoTo LOKAL_ERROR

  Dim osvi As OSVERSIONINFO

   osvi.dwOSVersionInfoSize = Len(osvi)
   If GetVersionEx(osvi) = 0 Then
      Exit Function
   End If
   GetWinPlatform = osvi.dwPlatformId
   
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "GetWinPlatform"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub AnalyseZusammenstellen(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL                As String
    Dim rsrs                As Recordset
    Dim siLug               As Single
    Dim dateLug             As Date
    Dim lKundenZR           As Long
    Dim lKundenVJZR         As Long
    Dim lWerktageZR         As Long
    Dim lWerktageVJZR       As Long
    Dim dNettoertragZR      As Double
    Dim dNettoertragVJZR    As Double
    Dim dNettoumsatzZR      As Double
    Dim dNettoUmsatzVJZR    As Double
    Dim dNettoRABATTZR      As Double
    Dim dNettoRABATTVJZR    As Double
    Dim dEinkaufZR          As Double
    Dim dEinkaufVJZR        As Double
    
    Dim lVon As Long
    Dim lBis As Long
    Dim lVonVJ As Long
    Dim lBisVJ As Long
    
    Dim sVon As String
    Dim sBis As String
    Dim svonVJ As String
    Dim sbisVJ As String
    Dim iAnzeigeZaehler As Integer
    Dim iBisZaehler As Integer
    Dim lGesbestand As Long
    Dim dLagersekwert As Double
    Dim lGesbestandPenner As Long
    Dim dLagersekwertPenner As Double
    
    Dim dumsatznull As Double
    Dim dumsatzKU As Double
    Dim lJahr As Long
    Dim bymonat As Byte
    Dim j As Integer
    Dim i As Integer
    Dim dProzent As Double
    Dim cZR As String
    Dim dUmsatzprokauf As Double
    Dim dTeileprokauf As Double
    Dim lNeukunden As Long
    Dim lNEINVERKAUF As Long
    Dim lArtikelschwund As Long
    Dim dArtikelschwundSEKWERT As Double
    Dim dArtikelschwundNettoErtrag As Double
    Dim sFilbez As String
    
    dumsatznull = 0
    dumsatzKU = 0
    
    iAnzeigeZaehler = 1
    iBisZaehler = 123
    
    If Month(DateValue(Now)) = 1 Then
        sVon = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
        sBis = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
    Else
        Select Case Month(DateValue(Now)) - 1
            Case 1, 3, 5, 7, 8, 10, 12
                sBis = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Case 2
                If Year(DateValue(Now)) = 2016 Then
                    sBis = "29.02.2016"
                ElseIf Year(DateValue(Now)) = 2020 Then
                    sBis = "29.02.2020"
                ElseIf Year(DateValue(Now)) = 2024 Then
                    sBis = "29.02.2024"
                Else
                    sBis = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End If
            Case Else
                sBis = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
        End Select
        
        sVon = Format(DateValue(sBis) + 1, "DD.MM.YYYY")
        sVon = Format(Day(DateValue(sVon)) & "." & Month(DateValue(sVon)) & "." & Year(DateValue(sVon)) - 1, "DD.MM.YYYY")
        
    End If

    svonVJ = Left(sVon, 6) & CInt(Right(sVon, 4)) - 1
    sbisVJ = Left(sBis, 6) & CInt(Right(sBis, 4)) - 1
    
    If sbisVJ = "29.02.2011" Or sbisVJ = "29.02.11" Then
        sbisVJ = "28.02.2011"
    ElseIf sbisVJ = "29.02.2015" Or sbisVJ = "29.02.15" Then
        sbisVJ = "28.02.2015"
    ElseIf sbisVJ = "29.02.2019" Or sbisVJ = "29.02.19" Then
        sbisVJ = "28.02.2019"
    ElseIf sbisVJ = "29.02.2023" Or sbisVJ = "29.02.23" Then
        sbisVJ = "28.02.2023"
    End If
    
    lVon = CLng(DateValue(sVon))
    lBis = CLng(DateValue(sBis))
    
    lVonVJ = CLng(DateValue(svonVJ))
    lBisVJ = CLng(DateValue(sbisVJ))
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    Screen.MousePointer = 11
    
    'Lagerumschlag Felder: lugakt , lugaktdate
    
    siLug = ermavgLUG
    dateLug = ermDateLUG
    
    sSQL = "Delete from GANALYSE "
    gdBase.Execute sSQL, dbFailOnError
    
    If CInt(gcFilNr) > 0 Then
        sFilbez = ermFilBez(CInt(gcFilNr))
    Else
        sFilbez = ""
    End If
    
    sSQL = "Insert into GANALYSE (Datum,LUGAKT,LUGAKTDATE,Filiale,FilialeBEZ,SENDOK) values "
    sSQL = sSQL & "('" & DateValue(Now) & "','" & siLug & "','" & dateLug & "'," & CInt(gcFilNr) & ",'" & sFilbez & "',False) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'Kundenfrequenz
    
    'Kundenzahl Vorjahresbereich
    lKundenZR = ermkunz8(lVon, lBis)
    lKundenVJZR = ermkunz8(lVonVJ, lBisVJ)
    
    lWerktageZR = ermWerktagemalFilialen(lVon, lBis)
    lWerktageVJZR = ermWerktagemalFilialen(lVonVJ, lBisVJ)
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET KundenfrequenzZR = " & lKundenZR
    sSQL = sSQL & ", KundenfrequenzVJZR = " & lKundenVJZR
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET KundenfrequenzENTWabs = KundenfrequenzZR - KundenfrequenzVJZR "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET KundenfrequenzENTWrela = 100 * KundenfrequenzENTWabs /KundenfrequenzZR "
    sSQL = sSQL & " where KundenfrequenzZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET ZRvon = " & lVon
    sSQL = sSQL & ", ZRbis = " & lBis
    sSQL = sSQL & ", VJZRvon = " & lVonVJ
    sSQL = sSQL & ", VJZRbis = " & lBisVJ
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET werktageZR = " & lWerktageZR
    sSQL = sSQL & ", werktageVJZR = " & lWerktageVJZR
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET KundenProTagproFilZR = KundenfrequenzZR /werktageZR "
    sSQL = sSQL & " where werktageZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET KundenProTagproFilVJZR = KundenfrequenzVJZR /werktageVJZR "
    sSQL = sSQL & " where werktageVJZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'nettoumsatz
    
    dNettoumsatzZR = ermNettoumsatz(lVon, lBis)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    dNettoRABATTZR = ermNettorabatt(lVon, lBis)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    dEinkaufZR = ermEinkauf(lVon, lBis)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    dNettoUmsatzVJZR = ermNettoumsatz(lVonVJ, lBisVJ)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    dNettoRABATTVJZR = ermNettorabatt(lVonVJ, lBisVJ)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    dEinkaufVJZR = ermEinkauf(lVonVJ, lBisVJ)
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET NettoUmsatzZR = '" & dNettoumsatzZR & "'"
    sSQL = sSQL & ", NettoUmsatzVJZR = '" & dNettoUmsatzVJZR & "'"
    sSQL = sSQL & ", NettoEinkaufZR = '" & dEinkaufZR & "'"
    sSQL = sSQL & ", NettoEinkaufVJZR = '" & dEinkaufVJZR & "'"
    sSQL = sSQL & ", NettoRabattZR = '" & dNettoRABATTZR & "'"
    sSQL = sSQL & ", NettoRabattVJZR = '" & dNettoRABATTVJZR & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'Nettoertrag
    
    dNettoertragZR = dNettoumsatzZR - dEinkaufZR
    dNettoertragVJZR = dNettoUmsatzVJZR - dEinkaufVJZR
    
    sSQL = "Update GANALYSE SET NettoertragZR = '" & dNettoertragZR & "'"
    sSQL = sSQL & ", NettoertragVJZR = '" & dNettoertragVJZR & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET NettoertragProTagproFilZR = NettoertragZR /werktageZR "
    sSQL = sSQL & " where werktageZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET NettoertragProTagproFilVJZR = NettoertragVJZR /werktageVJZR "
    sSQL = sSQL & " where werktageVJZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET NettospanneZR = NettoertragZR *100 /NettoUmsatzZR "
    sSQL = sSQL & " where NettoUmsatzZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Update GANALYSE SET NettospannevjZR = NettoertragvjZR *100 /NettoUmsatzvjZR "
    sSQL = sSQL & " where NettoUmsatzvjZR <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'Lagerbestand und SEKWERT
    
    lGesbestand = ermLagerbestand
    dLagersekwert = ermlagersekwert
    
    sSQL = "Update GANALYSE SET BESTAND = " & lGesbestand & " "
    sSQL = sSQL & ", LagerwertSEK = '" & dLagersekwert & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'LagerPennerbestand und Penner - SEKWERT
    
    lGesbestandPenner = ermLagerbestandPenner
    dLagersekwertPenner = ermlagersekwertPenner
    
    sSQL = "Update GANALYSE SET PennerBESTAND = " & lGesbestandPenner & " "
    sSQL = sSQL & ", PennerLagerwertSEK = '" & dLagersekwertPenner & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    'Kundenanteile
    'Umsatzanteil in Prozent
    
    loeschNEW "KUNZTohneK", gdBase
    
    sSQL = "Select * into KUNZTohneK from Kassjour where  "
    sSQL = sSQL & " Kundnr = 0 "
    sSQL = sSQL & " and adate > " & CLng(DateValue(Now) - 124)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    CheckIndex "KUNZTohneK", "adate", "", gdBase
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTohneK", "Filiale", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTohneK", "Kundnr", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    loeschNEW "KUNZTmitK", gdBase
    
    sSQL = "Select * into KUNZTmitK from Kassjour where  "
    sSQL = sSQL & " Kundnr > 0 "
    sSQL = sSQL & " and adate > " & CLng(DateValue(Now) - 124)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    CheckIndex "KUNZTmitK", "adate", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTmitK", "Filiale", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTmitK", "Kundnr", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    
    bymonat = Month(Now)
    lJahr = Year(Now)
    
        
    For j = 1 To 3
        bymonat = bymonat - 1
        If bymonat = 0 Then
            bymonat = 12
            lJahr = lJahr - 1
        End If
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        dumsatznull = ermKundenumsatz(False, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        dumsatzKU = ermKundenumsatz(True, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        lArtikelschwund = ermArtikelSchwundZR(bymonat, lJahr)
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dArtikelschwundSEKWERT = ermArtikelSchwundSEKWERTZR(bymonat, lJahr)
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dArtikelschwundNettoErtrag = ermArtikelSchwundNettoertragZR(bymonat, lJahr)
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        lUMSNullkunden = 0
        lUMSKUkunden = 0
        
        lUMSNullkunden = ermKundenumsatzproKauf(False, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        lUMSKUkunden = ermKundenumsatzproKauf(True, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    
        lNeukunden = ermNeuKundenZR(bymonat, lJahr, CInt(gcFilNr))
        lNEINVERKAUF = ermNeinVKZR(bymonat, lJahr)
        
        cZR = bymonat & "/" & lJahr
        
        sSQL = "Update GANALYSE SET KBZR" & j & " = '" & cZR & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        sSQL = "Update GANALYSE SET ARTSCHWUND" & j & " = " & lArtikelschwund
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update GANALYSE SET ARTSCHWUNDSEKW" & j & " = '" & dArtikelschwundSEKWERT & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update GANALYSE SET ARTSCHWUNDNERTRAG" & j & " = '" & dArtikelschwundNettoErtrag & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        dProzent = 0
        If (dumsatznull + dumsatzKU) <> 0 Then
            dProzent = (100 * dumsatznull) / (dumsatznull + dumsatzKU)
        End If
        
        sSQL = "Update GANALYSE SET KBNULL" & j & " = '" & dProzent & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dProzent = 0
        If (dumsatznull + dumsatzKU) <> 0 Then
            dProzent = (100 * dumsatzKU) / (dumsatznull + dumsatzKU)
        End If
        
        sSQL = "Update GANALYSE SET KBmit" & j & " = '" & dProzent & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dUmsatzprokauf = 0
        If lUMSNullkunden <> 0 Then
            dUmsatzprokauf = dumsatznull / lUMSNullkunden
        End If
        
        sSQL = "Update GANALYSE SET TproKnull" & j & " = '" & dUmsatzprokauf & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dUmsatzprokauf = 0
        If lUMSKUkunden <> 0 Then
            dUmsatzprokauf = dumsatzKU / lUMSKUkunden
        End If
        
        sSQL = "Update GANALYSE SET TproKmit" & j & " = '" & dUmsatzprokauf & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        sSQL = "Update GANALYSE SET NKU" & j & " = " & lNeukunden
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        sSQL = "Update GANALYSE SET NVK" & j & " = " & lNEINVERKAUF
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
    Next j
    
    'vorjahr
    
    loeschNEW "KUNZTohneK", gdBase
    
    sSQL = "Select * into KUNZTohneK from Kassjour where  "
    sSQL = sSQL & " Kundnr = 0 "
    sSQL = sSQL & " and adate between " & CLng(DateValue(Now) - 124 - 364) & " and " & CLng(DateValue(Now) - 364)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    CheckIndex "KUNZTohneK", "adate", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTohneK", "Filiale", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTohneK", "Kundnr", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    loeschNEW "KUNZTmitK", gdBase
    
    sSQL = "Select * into KUNZTmitK from Kassjour where  "
    sSQL = sSQL & " Kundnr > 0 "
    sSQL = sSQL & " and adate between " & CLng(DateValue(Now) - 124 - 364) & " and " & CLng(DateValue(Now) - 364)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    CheckIndex "KUNZTmitK", "adate", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTmitK", "Filiale", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    CheckIndex "KUNZTmitK", "Kundnr", "", gdBase
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    Dim dumsatznullVJ As Double
    Dim dumsatzKUVJ As Double
    Dim lUMSKUkundenVJ As Long
    Dim lUMSNullkundenVJ As Long
    
    dumsatznullVJ = 0
    dumsatzKUVJ = 0
    
    
    bymonat = Month(Now)
    lJahr = Year(Now) - 1
        
    For j = 1 To 3
        bymonat = bymonat - 1
        If bymonat = 0 Then
            bymonat = 12
            lJahr = lJahr - 1
        End If
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        dumsatznullVJ = ermKundenumsatz(False, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        dumsatzKUVJ = ermKundenumsatz(True, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        lUMSNullkundenVJ = 0
        lUMSKUkundenVJ = 0
    
        lUMSNullkundenVJ = ermKundenumsatzproKauf(False, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        lUMSKUkundenVJ = ermKundenumsatzproKauf(True, bymonat, lJahr, CInt(gcFilNr))
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
        lNeukunden = ermNeuKundenZR(bymonat, lJahr, CInt(gcFilNr))
        lNEINVERKAUF = ermNeinVKZR(bymonat, lJahr)
        
        cZR = bymonat & "/" & lJahr
        
        sSQL = "Update GANALYSE SET KBZRVJ" & j & " = '" & cZR & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dProzent = 0
        If (dumsatznullVJ + dumsatzKUVJ) <> 0 Then
            dProzent = (100 * dumsatznullVJ) / (dumsatznullVJ + dumsatzKUVJ)
        End If
        
        sSQL = "Update GANALYSE SET KBNULLvj" & j & " = '" & dProzent & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dProzent = 0
        If (dumsatznullVJ + dumsatzKUVJ) <> 0 Then
            dProzent = (100 * dumsatzKUVJ) / (dumsatznullVJ + dumsatzKUVJ)
        End If
        
        sSQL = "Update GANALYSE SET KBmitvj" & j & " = '" & dProzent & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dUmsatzprokauf = 0
        If lUMSNullkundenVJ <> 0 Then
            dUmsatzprokauf = dumsatznullVJ / lUMSNullkundenVJ
        End If
        
        sSQL = "Update GANALYSE SET TproKnullvj" & j & " = '" & dUmsatzprokauf & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        dUmsatzprokauf = 0
        If lUMSKUkundenVJ <> 0 Then
            dUmsatzprokauf = dumsatzKUVJ / lUMSKUkundenVJ
        End If
        
        sSQL = "Update GANALYSE SET TproKmitvj" & j & " = '" & dUmsatzprokauf & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        sSQL = "Update GANALYSE SET NKUvj" & j & " = " & lNeukunden
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
        sSQL = "Update GANALYSE SET NVKvj" & j & " = " & lNEINVERKAUF
        gdBase.Execute sSQL, dbFailOnError
        
        anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
        
    Next
    
    loeschNEW "KUNZTohneK", gdBase
    loeschNEW "KUNZTmitK", gdBase
    
    anzeige "black", iAnzeigeZaehler & " von " & iBisZaehler, lblanzeige: iAnzeigeZaehler = iAnzeigeZaehler + 1
    
    sSQL = "Delete from GANALYSEALL where datum =  " & CLng(DateValue(Now)) & "  and filiale = " & gcFilNr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into GANALYSEALL select * from GANALYSE "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "AnalyseZusammenstellen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Function GANALYSEAKTUELL() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    GANALYSEAKTUELL = False
    
    sSQL = "select top 1 DATUM from GANALYSEALL  "
    sSQL = sSQL & " where filiale =  " & gcFilNr
    sSQL = sSQL & " order by datum desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Datum) Then
            If Month(DateValue(Now)) = Month(rsrs!Datum) Then
                If Year(DateValue(Now)) = Year(rsrs!Datum) Then
                    GANALYSEAKTUELL = True
                End If
            End If
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "GANALYSEAKTUELL"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermkunz8(lVon As Long, lBis As Long) As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermkunz8 = 0
    
    cSQL = "Select sum(kunz1) as maxi from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermkunz8 = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermkunz8"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermWerktagemalFilialen(lVon As Long, lBis As Long) As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermWerktagemalFilialen = 0
    
    cSQL = "Select * from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        ermWerktagemalFilialen = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermWerktagemalFilialen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermFilBez(iFil As Integer) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermFilBez = ""
    
    cSQL = "Select * from Filialen  "
    cSQL = cSQL & " where filialnr =  " & iFil
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!FILIALNAME) Then
            ermFilBez = rsrs!FILIALNAME
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermFilBez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermEinkauf(lVon As Long, lBis As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermEinkauf = 0
    
    cSQL = "Select sum(ekpr1) as maxi from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermEinkauf = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermEinkauf"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNettoumsatz(lVon As Long, lBis As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    ermNettoumsatz = 0
    
    dWert = 0
    
    
    cSQL = "Select sum(umsv1) as maxi from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            
            ermNettoumsatz = (dWert * 100) / (100 + gdMWStV)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dWert = 0
    
    cSQL = "Select sum(umse1) as maxi from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            
            ermNettoumsatz = ermNettoumsatz + (dWert * 100) / (100 + gdMWStE)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dWert = 0
    
    cSQL = "Select sum(umso1) as maxi from umsatz "
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            
            ermNettoumsatz = ermNettoumsatz + (dWert * 100) / (100 + gdMWStO)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermNettoumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNettorabatt(lVon As Long, lBis As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    ermNettorabatt = 0
    
    dWert = 0
    
    cSQL = "Select sum((vkpr * Menge) -preis) as maxi from Kassjour "
    cSQL = cSQL & " where adate between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST = 'V' "
    cSQL = cSQL & " and abs(vkpr * Menge * 100) > abs(preis * 100) "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            ermNettorabatt = (dWert * 100) / (100 + gdMWStV)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dWert = 0
    
    cSQL = "Select sum((vkpr * Menge) -preis) as maxi from Kassjour "
    cSQL = cSQL & " where adate between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST = 'E' "
    cSQL = cSQL & " and abs(vkpr * Menge * 100) > abs(preis * 100) "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            ermNettorabatt = ermNettorabatt + (dWert * 100) / (100 + gdMWStE)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dWert = 0
    
    cSQL = "Select sum((vkpr * Menge) -preis) as maxi from Kassjour "
    cSQL = cSQL & " where adate between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST = 'O' "
    cSQL = cSQL & " and abs(vkpr * Menge * 100) > abs(preis * 100) "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dWert = rsrs!maxi
            ermNettorabatt = ermNettorabatt + dWert
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermNettorabatt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Lagerwerteschreiben(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset

    sSQL = " Select * from Lagerw where Datum = datevalue(now) "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 11
    
    loeschNEW "LAGERD", gdBase
    CreateTableT2 "LAGERD", gdBase
    
    sSQL = " Insert into LAGERD Select distinct(ARTNR) as artikelnummer,LINR,0 as LPZ,0 as EKPR,'' as Marke,0 as KVKPR1,'' as BEZEICH, BESTAND from Artikel where bestand > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "LAGERD", "artikelnummer", "", gdBase
    
    sSQL = "Update LAGERD inner join Artikel on LAGERD.artikelnummer = ARTIKEL.ARTNR "
    sSQL = sSQL & " Set LAGERD.LPZ = ARTIKEL.LPZ "
    sSQL = sSQL & " , LAGERD.EKPR = ARTIKEL.EKPR "
    sSQL = sSQL & " , LAGERD.BEZEICH = ARTIKEL.BEZEICH "
    sSQL = sSQL & " , LAGERD.KVKPR1 = ARTIKEL.KVKPR1 "
    gdBase.Execute sSQL, dbFailOnError
    

    'the same for GDPdU
    Speicher_Bestände_GDPdU
    'End GDPdU
    

    
    sSQL = "Update LAGERD inner join LINBEZ on LAGERD.LINR = LINBEZ.LINR  "
    sSQL = sSQL & " and LAGERD.LPZ = LINBEZ.LPZ "
    sSQL = sSQL & " set LAGERD.MARKE = LINBEZ.MARKE "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LagerMw Select Datevalue(now) as datum , Marke,sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERD "
    sSQL = sSQL & " group by Marke "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Lagerllw Select Datevalue(now) as datum , LINR,LPZ,sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERD "
    sSQL = sSQL & " group by LINR, LPZ "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Lagerlw Select Datevalue(now) as datum , LINR,sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERD "
    sSQL = sSQL & " group by LINR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Lagerw Select Datevalue(now) as datum , sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERD "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "LAGERD", gdBase
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "Lagerwerteschreiben"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
''Public Sub kundnr_zukartenNr(lblAnzeige As Label)
''On Error GoTo LOKAL_ERROR
''
''    Dim sSQL        As String
''    Dim rsrs        As Recordset
''    Dim cKundNr     As String
''
''    Screen.MousePointer = 11
''
''    anzeige "normal", "", lblAnzeige
''
''    sSQL = "Select kundnr ,kundkart from Kunden where kundnr > 100 "
''
''    Set rsrs = gdBase.OpenRecordset(sSQL)
''    If Not rsrs.EOF Then
''        rsrs.MoveFirst
''        Do While Not rsrs.EOF
''            If Not IsNull(rsrs!Kundnr) Then
''
''                cKundNr = rsrs!Kundnr
''
''                rsrs.Edit
''                rsrs!KUNDKART = fnMoveArtNr2EAN8_begin980(cKundNr)
''                rsrs.Update
''            End If
''            rsrs.MoveNext
''        Loop
''    End If
''    rsrs.Close
''
''
''
''    anzeige "normal", "", lblAnzeige
''
''    Screen.MousePointer = 0
''
''Exit Sub
''LOKAL_ERROR:
''
''    Fehler.gsDescr = err.Description
''    Fehler.gsNumber = err.Number
''    Fehler.gsFormular = "Modul8"
''    Fehler.gsFunktion = "kundnr_zukartenNr"
''    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
''
''    Fehlermeldung1
''
''End Sub
Public Sub SEKNUll(lblanzeige As Label, bBestand As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "", lblanzeige
    
    loeschNEW "SEKNULL", gdBase
    CreateTableT2 "SEKNULL", gdBase
    
    cSQL = "Insert into SEKNULL Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , LINR "
    
    cSQL = cSQL & " from ARTIKEL where EKPR <= 0 "
    cSQL = cSQL & " and gefuehrt = 'J'"
    
    If bBestand Then
        cSQL = cSQL & " and Bestand > 0 "
    End If

    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on SEKNULL(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SEKNULL inner join LISRT on SEKNULL.Linr = LISRT.Linr "
    cSQL = cSQL & " set SEKNULL.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    reportbildschirm "WKL024", "aWKL40ac"
    
    Pause (3)
    loeschNEW "SEKNULL", gdBase
    
    anzeige "normal", "", lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "SEKNUll"
    Fehler.gsFehlertext = "Beim Ermitteln ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SchnittEinkaufspreisbereinigung(lblanzeige As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "", lblanzeige
    
    'Hier werden alle SEK = 0 mit einem LEK > 0 gefüllt
    
    loeschNEW "SCHNITT_NULL", gdBase
    
    cSQL = "Create Table SCHNITT_NULL ( "
    cSQL = cSQL & " ARTNR Long "
    cSQL = cSQL & ", EKPR DOUBLE "
    cSQL = cSQL & ", LEKPR DOUBLE "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ARTIKEL "
    cSQL = cSQL & " set ekpr = 0 where Ekpr is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into SCHNITT_NULL Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , EKPR "
    cSQL = cSQL & " , 0 as LEKPR "
    cSQL = cSQL & " from ARTIKEL where EKPR <= 0 and KVKPR1 >= 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    'kleineste LEK-tabelle bauen
    
    loeschNEW "KL_ARTLIEF", gdBase
    
    cSQL = "Select Artnr, min(LEKPR) as MinLEKPR into KL_ARTLIEF"
    cSQL = cSQL & " from ARTLIEF where LEKPR > 0 group by Artnr  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SCHNITT_NULL inner join KL_ARTLIEF on SCHNITT_NULL.artnr = KL_ARTLIEF.Artnr"
    cSQL = cSQL & " set SCHNITT_NULL.Lekpr = KL_ARTLIEF.MinLEKPR "
    gdBase.Execute cSQL, dbFailOnError
    
    'Final updaten
    cSQL = "Update Artikel inner join SCHNITT_NULL on Artikel.artnr = SCHNITT_NULL.Artnr"
    cSQL = cSQL & " set Artikel.ekpr = SCHNITT_NULL.Lekpr "
    cSQL = cSQL & " where Artikel.EKPR = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    'Ende Teil 1
    
    
    cSQL = "Delete from SCHNITT_NULL where lekpr > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    'Teil 2
    
    SpalteAnfuegenNEW "SCHNITT_NULL", "verkauft", "BIT", gdBase
    
    sSQL = "Update SCHNITT_NULL inner join kassjour on SCHNITT_NULL.artnr = kassjour.artnr"
    sSQL = sSQL & " set SCHNITT_NULL.verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from SCHNITT_NULL where verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from SCHNITT_NULL where artnr in (Select artnr from Warengru) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index artnr on SCHNITT_NULL (artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from SCHNITT_NULL) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artikel where artnr in (Select artnr from SCHNITT_NULL) "
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Teil 2
    
    'bleiben immernoch welche übrig?
    loeschNEW "SCHNITT_NULL", gdBase
    
    cSQL = "Create Table SCHNITT_NULL ( "
    cSQL = cSQL & " ARTNR Long "
    cSQL = cSQL & ", EKPR DOUBLE "
    cSQL = cSQL & ", LEKPR DOUBLE "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into SCHNITT_NULL Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , EKPR "
    cSQL = cSQL & " , 0 as LEKPR "
    cSQL = cSQL & " from ARTIKEL where EKPR <= 0 and KVKPR1 >= 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    'Teil 3
    cSQL = "Update Kassjour k set K.EKPR = 0 "
    cSQL = cSQL & " where k.ekpr is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Kassjour k inner join Artikel a on k.artnr = a.artnr set K.EKPR = a.ekpr"
    cSQL = cSQL & " where k.ekpr = 0 "
    gdBase.Execute cSQL, dbFailOnError
    'Ende Teil 3
    
    
    loeschNEW "PREISE_FALSCH", gdBase
    
    cSQL = "Create Table PREISE_FALSCH ( "
    cSQL = cSQL & " ARTNR Long "
    cSQL = cSQL & ", EKPR DOUBLE "
    cSQL = cSQL & ", KVKPR1 DOUBLE "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into PREISE_FALSCH Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , EKPR "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " from ARTIKEL where 4 * kvkpr1 < ekpr and KVKPR1 >= 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    SpalteAnfuegenNEW "PREISE_FALSCH", "verkauft", "BIT", gdBase
    
    sSQL = "Update PREISE_FALSCH inner join kassjour on PREISE_FALSCH.artnr = kassjour.artnr"
    sSQL = sSQL & " set PREISE_FALSCH.verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from PREISE_FALSCH where verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from PREISE_FALSCH where artnr in (Select artnr from Warengru) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index artnr on PREISE_FALSCH (artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from PREISE_FALSCH) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artikel where artnr in (Select artnr from PREISE_FALSCH) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "PREISE_FALSCH", gdBase
    
    cSQL = "Create Table PREISE_FALSCH ( "
    cSQL = cSQL & " ARTNR Long "
    cSQL = cSQL & ", EKPR DOUBLE "
    cSQL = cSQL & ", KVKPR1 DOUBLE "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into PREISE_FALSCH Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , EKPR "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " from ARTIKEL where 4 * kvkpr1 < ekpr and KVKPR1 >= 0  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    reportbildschirm "", "aWKL33k"
    reportbildschirm "", "aWKL33m"
    
    
'    select * from artikel where 4 * kvkpr1 < ekpr
    Pause (2)
    loeschNEW "SCHNITT_NULL", gdBase
    

    
    anzeige "normal", "Fertig", lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "SchnittEinkaufspreisbereinigung"
    Fehler.gsFehlertext = "Beim Ermitteln ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SchnittEinkaufspreisbereinigungT2(lblanzeige As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    anzeige "normal", "", lblanzeige
    
    If NewTableSuchenDBKombi("PREISE_FALSCH", gdBase) Then
    
        cSQL = "Update PREISE_FALSCH k inner join Artikel a on k.artnr = a.artnr set K.EKPR = a.ekpr"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Kassjour k inner join PREISE_FALSCH a on k.artnr = a.artnr set K.EKPR = a.ekpr"
        gdBase.Execute cSQL, dbFailOnError
        
        anzeige "normal", "Fertig", lblanzeige
    Else
        anzeige "rot", "Abbruch", lblanzeige
    End If
    
    loeschNEW "PREISE_FALSCH", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "SchnittEinkaufspreisbereinigungT2"
    Fehler.gsFehlertext = "Beim Ermitteln ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub ArtEanerstellen(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim sPfad       As String
    
    Screen.MousePointer = 11
    
    loeschNEW "ARTEAN", gdBase
    CreateTableT2 "ARTEAN", gdBase
    
    anzeige "normal", "EAN", lblanzeige
    
    sSQL = " Insert into ARTEAN Select ARTNR,EAN from Artikel "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "EAN2", lblanzeige
    
    sSQL = " Create index Ean on ARTEAN(ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into ARTEAN Select ARTNR,EAN2 as EAN from Artikel "
'    sSQL = sSQL & " where ean2 not in (Select ean from ARTEAN) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "EAN3", lblanzeige
    
    sSQL = " Insert into ARTEAN Select ARTNR,EAN3 as EAN from Artikel "
'    sSQL = sSQL & " where ean3 not in (Select ean from ARTEAN) "
    gdBase.Execute sSQL, dbFailOnError
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    Kill sPfad & "BOX\ARTEAN.DBF"

    sSQL = "Select * into ARTEAN IN '" & sPfad & "BOX" & "' 'dbase IV;' from ARTEAN"
    gdBase.Execute sSQL, dbFailOnError
    
    MsgBox "Die Datei befindet sich hier: " & sPfad & "BOX\ARTEAN.DBF", vbInformation, "Winkiss Hinweis:"
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul8"
        Fehler.gsFunktion = "ArtEanerstellen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub bonusfArtikel(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    Screen.MousePointer = 11

    anzeige "normal", "Bonus = ja", lblanzeige
    
    sSQL = "Update Artikel set bonus_ok = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Bonus = nein", lblanzeige
    
    sSQL = "Update Artikel inner join Warengru on Artikel.artnr = Warengru.artnr  set bonus_ok = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "bonusfArtikel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermLagerbestandPenner() As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLagerbestandPenner = 0
    
    cSQL = "select top 1 Datum,Best from Penlagerw order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!best) Then
            ermLagerbestandPenner = rsrs!best
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermLagerbestandPenner"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermlagersekwertPenner() As Single
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermlagersekwertPenner = 0
    
    cSQL = "select top 1 Datum, SEK from PENlagerw order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sEK) Then
            ermlagersekwertPenner = rsrs!sEK
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermlagersekwertPenner"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermlagersekwert() As Single
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermlagersekwert = 0
    
    cSQL = "select TOP 1 Datum, SEK from lagerw order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sEK) Then
            ermlagersekwert = rsrs!sEK
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermlagersekwert"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLagerbestand() As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLagerbestand = 0
    
    cSQL = "select top 1 Datum,Best from lagerw order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!best) Then
            ermLagerbestand = rsrs!best
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermLagerbestand"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub PennerBestundSEK(lblanzeige As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    Dim lcount As Long
    
    sSQL = " Select * from PENLagerw where Datum = datevalue(now) "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    loeschNEW "ART69", gdBase
    CreateTableT2 "ART69", gdBase

    sSQL = " Insert into ART69 select  ARTNR"
    sSQL = sSQL & " , EKPR "
    sSQL = sSQL & " , BESTAND "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , '' as Marke "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & "  from Artikel "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART69 inner join LINBEZ on ART69.LINR = LINBEZ.LINR  "
    sSQL = sSQL & " and ART69.LPZ = LINBEZ.LPZ "
    sSQL = sSQL & " set ART69.MARKE = LINBEZ.MARKE "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "neue Artikel ausschließen", lblanzeige
    
    sSQL = "Delete from ART69 where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ART69 where bestand is null "
    gdBase.Execute sSQL, dbFailOnError
    
    If CInt(gcFilNr) = 0 Then
    
    Else
        If NewTableSuchenDBKombi("UMLAGER", gdBase) Then
            CheckIndex "Umlager", "adate", "", gdBase
            CheckIndex "Umlager", "artnr", "", gdBase
        End If
    End If
    
    Set rsrs = gdBase.OpenRecordset("ART69")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
            
                anzeige "normal", lcount & " Artikel noch...", lblanzeige
                lcount = lcount - 1
                
                rsrs.Edit
                rsrs!ERSTDAT = ErmFirstZugang(rsrs!artnr)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "neue Artikel entfernen", lblanzeige
    
    sSQL = " delete from ART69  where ERSTDAT > datevalue(now) - 180 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "maxKass", gdBase
    
    sSQL = "Select max(adate) as maxdate,artnr  into maxkass from Kassjour group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "maxKass", "Artnr", "", gdBase
    
    loeschNEW "maxKast", gdBase
    
    sSQL = "Select * into maxKast from maxKass where artnr in (Select artnr from ART69) "
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "maxKast", "Artnr", "", gdBase
    loeschNEW "maxKass", gdBase
    
    anzeige "normal", "relevante Artikel ermitteln", lblanzeige
    
    Set rsrs = gdBase.OpenRecordset("ART69")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                ldifferenz = 0
                rsrs.Edit
                datLVK = ErmlzVK1(cART, gdBase)
                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                Select Case ldifferenz
           
                    Case Is > 731
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 24 Monaten"
                        End If
                        
                    Case Is > 701
                        ctmp = "seit 23 Monaten"
                    Case Is > 671
                        ctmp = "seit 22 Monaten"
                    Case Is > 640
                        ctmp = "seit 21 Monaten"
                    Case Is > 610
                        ctmp = "seit 20 Monaten"
                    Case Is > 579
                        ctmp = "seit 19 Monaten"
                    Case Is > 549
                        ctmp = "seit 18 Monaten"
                    Case Is > 519
                        ctmp = "seit 17 Monaten"
                    Case Is > 488
                        ctmp = "seit 16 Monaten"
                    Case Is > 457
                        ctmp = "seit 15 Monaten"
                    Case Is > 426
                        ctmp = "seit 14 Monaten"
                    Case Is > 395
                        ctmp = "seit 13 Monaten"
                    Case Is > 365
                        ctmp = "seit 12 Monaten"
                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close
    
    anzeige "normal", "nur Penner schreiben", lblanzeige
    
    loeschNEW "maxKast", gdBase
    loeschNEW "maxZUt", gdBase
    
    sSQL = "Delete from ART69 where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ART69 where Monat is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into PENLagermw Select Datevalue(now) as datum, marke , sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from ART69 group by marke "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into PENLagerLw Select Datevalue(now) as datum ,linr, sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from ART69 group by linr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into PENLagerLLw Select Datevalue(now) as datum ,linr,lpz, sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from ART69 group by linr,lpz "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into PENLagerw Select Datevalue(now) as datum , sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from ART69 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART69", gdBase
    
    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "PennerBestundSEK"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function ErmFirstZugang(cART As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmFirstZugang = DateValue("01.01.1980")
    
    Dim cSQL As String
    Dim rsINB As Recordset

        cSQL = "Select min(adate) as mindate from Zugang where ARTNR = " & cART & " "

    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MinDate) Then
            ErmFirstZugang = rsINB!MinDate
        End If
    End If
    rsINB.Close: Set rsINB = Nothing
    
    
    'check auch umlager
    
    If CInt(gcFilNr) > 0 Then
    
        Dim FirstUmlager As Date
        FirstUmlager = DateValue("01.01.1980")
    
        cSQL = "Select min(adate) as mindate from Umlager where ARTNR = " & cART & " "
        Set rsINB = gdBase.OpenRecordset(cSQL)
        If Not rsINB.EOF Then
            If Not IsNull(rsINB!MinDate) Then
                FirstUmlager = rsINB!MinDate
            End If
        End If
        rsINB.Close: Set rsINB = Nothing
        
        If FirstUmlager <> DateValue("01.01.1980") Then
            If FirstUmlager < ErmFirstZugang Then
                ErmFirstZugang = FirstUmlager
            End If
        End If
        
        If ErmFirstZugang = DateValue("01.01.1980") And FirstUmlager <> DateValue("01.01.1980") Then
            
            ErmFirstZugang = FirstUmlager
            
        End If
        
        
        
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmFirstZugang"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function

Public Function ErmlzVK1(cART As String, db As Database) As String
    On Error GoTo LOKAL_ERROR
    
    ErmlzVK1 = "0"
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select  maxdate from maxKast where ARTNR = " & cART & "  "
    Set rsINB = db.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVK1 = rsINB!MaxDate
        
        End If
    
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmlzVK1"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermavgLUG() As Single
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermavgLUG = 0
    
    cSQL = "select top 1 Datum,avgLUG from LUGEVER order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!avgLUG) Then
            ermavgLUG = rsrs!avgLUG
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermavgLUG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermDateLUG() As Date
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermDateLUG = 0
    
    cSQL = "select top 1 Datum from LUGEVER order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Datum) Then
            ermDateLUG = rsrs!Datum
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "ermDateLUG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

