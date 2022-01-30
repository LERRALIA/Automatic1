Attribute VB_Name = "Modul4"
  Option Explicit
  
  
  
Public Declare Sub GetSystemTime Lib "kernel32" ( _
  lpSystemTime As SYSTEMTIME)
 
Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type



' zunächst die benötigte API-Funktion
Private Declare Function GetKeyboardState Lib "user32" ( _
  pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" ( _
  lppbKeyState As Byte) As Long
 



Public Declare Function GetKeyState Lib "user32" ( _
  ByVal nVirtKey As Long) As Integer

  
Private Declare Sub keybd_event Lib "user32" ( _
  ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)
 
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_KEYUP = &H2
  
    Public Const NORMAL_PRIORITY_CLASS      As Long = &H20&
    Public Const STATUS_WAIT_0              As Long = &H0
    Public Const WAIT_OBJECT_0              As Long = STATUS_WAIT_0
  
    Private Declare Function InputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
  
    Public Type STARTUPINFO
      cb              As Long
      lpReserved      As Long
      lpDesktop       As Long
      lpTitle         As Long
      dwX             As Long
      dwY             As Long
      dwXSize         As Long
      dwYSize         As Long
      dwXCountChars   As Long
      dwYCountChars   As Long
      dwFillAttribute As Long
      dwFlags         As Long
      wShowWindow     As Integer
      cbReserved2     As Integer
      lpReserved2     As Long
      hStdInput       As Long
      hStdOutput      As Long
      hStdError       As Long
    End Type
  
  Public Type PROCESS_INFORMATION
      hProcess    As Long
      hThread     As Long
      dwProcessID As Long
      dwThreadID  As Long
  End Type
  ' Diese Routine schaltet CAPS-Lock ein,
' falls ausgeschaltet und umgekehrt
Public Sub KeyboardChangeState(ByVal lpKey As Long)
  ReDim kBuffer(256) As Byte
  GetKeyboardState kBuffer(0)
  If kBuffer(lpKey) And 1 Then
    kBuffer(lpKey) = 0
  Else
    kBuffer(lpKey) = 1
  End If
  SetKeyboardState kBuffer(0)
End Sub

  

  Public Sub DrawPiePiece(lcolor As Long, ByRef fStart As Double, ByRef fEnd As Double, PicX As PictureBox, siRadius As Single)
        Const PI            As Double = 3.14159265359
        Const CircleEnd     As Double = -2 * PI
        Dim dStart          As Double
        Dim dEnd            As Double
        PicX.FillColor = lcolor
        PicX.FillStyle = 0
        dStart = fStart * (CircleEnd / 100)
        dEnd = fEnd * (CircleEnd / 100)
        
        PicX.Circle (siRadius, siRadius), siRadius, , dStart, dEnd
    End Sub
  Public Sub LeseLIZENZ()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cPfad As String
    
    Screen.MousePointer = 11
    
    gbLizenz = False
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iFileNr = FreeFile
    Open cPfad & "ZNEZIL.CFG" For Binary As #iFileNr
'    Open cPfad & "LIZENZ.CFG" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        
        ctmp = SwapStr(ctmp, Chr(13), "")
        ctmp = SwapStr(ctmp, Chr(10), "")
        If Trim(ctmp) = "Stammdaten 2008" Then
            gbLizenz = True
        End If
        Close iFileNr
    Else
        Close iFileNr
    End If
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 76 Or err.Number = 52 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "LeseLIZENZ"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub LeseLIZENZ_INDI()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cPfad As String
    
    Screen.MousePointer = 11
    
    gbLizenzINDI = False
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iFileNr = FreeFile
    Open cPfad & "ZNEZILINDI.CFG" For Binary As #iFileNr

    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        
        ctmp = SwapStr(ctmp, Chr(13), "")
        ctmp = SwapStr(ctmp, Chr(10), "")
        If Trim(ctmp) = "Stammdaten 2008" Then
            gbLizenzINDI = True
        End If
        Close iFileNr
    Else
        Close iFileNr
    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 76 Or err.Number = 52 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "LeseLIZENZ_INDI"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub DruckeGrundPreisEtikettenWKL30kleinspezial(acArtNr() As String, lAnzahl As Long, srepname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    
    Dim iFileNr As Integer

    loeschNEW "DRU_GRUN", gdBase
    
    cSQL = "Create Table DRU_GRUN ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "

    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")
    
    cSQL = "Select * from DRU_GRUN"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    rsZiel.AddNew
    rsZiel!FirmaName = ""
    rsZiel!artnr = Null
    rsZiel!BEZEICH = ""
    rsZiel!Barcode = ""
    rsZiel!LIBESNR = ""
    rsZiel!vkpr = Null
    rsZiel!vkpr_EUR = Null
    rsZiel!INHALT = Null
    rsZiel!INHALTBEZ = ""
    rsZiel!DRUCKDATUM = ""
    rsZiel!GRUNDPREIS = ""
    rsZiel!GRUND_INH = Null
    rsZiel!GRUND_DM = Null
    rsZiel!GRUND_EUR = Null
    rsZiel.Update
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If rsrs.EOF Then

        Else
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich
            
            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            rsZiel!EAN = cEAN
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!LIBESNR = rsrs!LIBESNR
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If
                
            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If
                
                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            '//DRU_GRUN
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    loeschNEW "DRU_GRUN", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "DRU_GRUN"
    
    cSQL = "Delete from DRU_GRUN where artnr is null"
    gdApp.Execute cSQL, dbFailOnError

    reportbildschirmToPrinterAPP srepname, gcEtikettenDrucker
    
    
      
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "DruckeGrundPreisEtikettenWKL30kleinspezial"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub DruckeTLPRegaletikett40x25Variante1(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

        
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            rsZiel!EAN1 = rsrs!EAN 'cEAN

            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode

            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett40x25Variante4(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

        
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            

            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25Variante4"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett40x25Variante5(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            
            'guck doch mal ob der EAN 13stellig ist
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
                If Len(cEAN13) < 13 Then
                    cEAN13 = ""
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                    If Len(cEAN13) < 13 Then
                        cEAN13 = ""
                    End If
                    
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                    
                    If Len(cEAN13) < 13 Then
                        cEAN13 = ""
                    End If
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            
            
            'im ersten Durchlauf war nichts zu machen - keine 13er
            'jetzt auf 8er prüfen
            
            If cEAN13 = "" Then
            
                cEAN13 = ""
                
                If Not IsNull(rsrs!EAN) Then
                    cEAN13 = rsrs!EAN 'cEAN13
                    If Len(cEAN13) < 8 Then
                        cEAN13 = ""
                    End If
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN2) Then
                        cEAN13 = rsrs!EAN2 'cEAN13
                        If Len(cEAN13) < 8 Then
                            cEAN13 = ""
                        End If
                    End If
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN3) Then
                        cEAN13 = rsrs!EAN3 'cEAN13
                        If Len(cEAN13) < 8 Then
                            cEAN13 = ""
                        End If
                    End If
                End If
                
                cEAN13 = Trim(cEAN13)
            End If
            
            '2.Durchlauf Ende
            
            
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
'            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            
            Dim sEtikettKürzel As String
            Dim sEtikettLibesnr As String
            
            If glEtiExArtikel_linr <> 0 Then
                sEtikettKürzel = ermLiefKürzelmitLiefvorgabe(rsrs!artnr, glEtiExArtikel_linr)
                sEtikettLibesnr = ermLiefLIBESNRmitLiefvorgabe(rsrs!artnr, glEtiExArtikel_linr)
            Else
                sEtikettKürzel = ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase)
                sEtikettLibesnr = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
            
            End If
                        
            
            
            'wenn Ex dann auch beim Lief bleiben
            rsZiel!DRUCKDATUM = Left(sEtikettKürzel, 3) & " " & cDruckdatum
            rsZiel!LIBESNR = sEtikettLibesnr
            
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25Variante5"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett40x25Variante7(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
   

    cDruckdatum = "KW" & DatePart("ww", DateValue(Now))

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            
            'guck doch mal ob der EAN 13stellig ist
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
                If Len(cEAN13) < 13 Then
                    cEAN13 = ""
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                    If Len(cEAN13) < 13 Then
                        cEAN13 = ""
                    End If
                    
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                    
                    If Len(cEAN13) < 13 Then
                        cEAN13 = ""
                    End If
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            
            
            'im ersten Durchlauf war nichts zu machen - keine 13er
            'jetzt auf 8er prüfen
            
            If cEAN13 = "" Then
            
                cEAN13 = ""
                
                If Not IsNull(rsrs!EAN) Then
                    cEAN13 = rsrs!EAN 'cEAN13
                    If Len(cEAN13) < 8 Then
                        cEAN13 = ""
                    End If
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN2) Then
                        cEAN13 = rsrs!EAN2 'cEAN13
                        If Len(cEAN13) < 8 Then
                            cEAN13 = ""
                        End If
                    End If
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN3) Then
                        cEAN13 = rsrs!EAN3 'cEAN13
                        If Len(cEAN13) < 8 Then
                            cEAN13 = ""
                        End If
                    End If
                End If
                
                cEAN13 = Trim(cEAN13)
            End If
            
            '2.Durchlauf Ende
            
            
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
'            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            
            Dim sEtikettKürzel As String
            Dim sEtikettLibesnr As String
            
            If glEtiExArtikel_linr <> 0 Then
                sEtikettKürzel = ermLiefKürzelmitLiefvorgabe(rsrs!artnr, glEtiExArtikel_linr)
                sEtikettLibesnr = ermLiefLIBESNRmitLiefvorgabe(rsrs!artnr, glEtiExArtikel_linr)
            Else
                sEtikettKürzel = ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase)
                sEtikettLibesnr = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
            
            End If
                        
            
            
            'wenn Ex dann auch beim Lief bleiben
            rsZiel!DRUCKDATUM = Left(sEtikettKürzel, 3) & " " & cDruckdatum
            rsZiel!LIBESNR = sEtikettLibesnr
            
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25Variante7"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett40x25Var_Dronova(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cMWST As String

    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(50)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            
            cMWST = "V"
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
                If Len(cEAN13) <> 13 And Len(cEAN13) <> 8 Then
                    cEAN13 = ""
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                    If Len(cEAN13) <> 13 And Len(cEAN13) <> 8 Then
                        cEAN13 = ""
                    End If
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                    If Len(cEAN13) <> 13 And Len(cEAN13) <> 8 Then
                        cEAN13 = ""
                    End If
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            
            'im ersten Durchlauf war nichts zu machen - keine 13er oder 8er
            'jetzt irgendwas

            If cEAN13 = "" Then

                cEAN13 = ""

                If Not IsNull(rsrs!EAN) Then
                    cEAN13 = rsrs!EAN
                End If

                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN2) Then
                        cEAN13 = rsrs!EAN2
                    End If
                End If

                If cEAN13 = "" Then
                    If Not IsNull(rsrs!EAN3) Then
                        cEAN13 = rsrs!EAN3 'cEAN13
                    End If
                End If

                cEAN13 = Trim(cEAN13)
            End If

            '2.Durchlauf Ende
            
            
            
            
            
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
        
            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            

            Dim sEtikettKürzel As String
            Dim sEtikettLibesnr As String
            Dim sEtikettVpe As String
            Dim sEtikettHS As String
            
            Dim bLieferant_ist_enthalten As Boolean
            
            Dim lLinr As Long
            bLieferant_ist_enthalten = False
            
            
            If frmWKL30.Label1(3).Caption <> "undefiniert" Then
                lLinr = Val(frmWKL30.Label1(3).Caption)
                If gibtesdiesenLieferantenhier(rsrs!artnr, lLinr) Then
                    bLieferant_ist_enthalten = True
                    
                    sEtikettKürzel = ermLiefKürzelmitLiefvorgabe(rsrs!artnr, lLinr)
                    sEtikettLibesnr = ermLiefLIBESNRmitLiefvorgabe(rsrs!artnr, lLinr)
                    sEtikettVpe = ermLief_VPE_mitLiefvorgabe(rsrs!artnr, lLinr)
                    sEtikettHS = ermLief_Handelsspanne_mitLiefvorgabe(rsrs!artnr, CStr(dVkPr), cMWST, lLinr)
                End If
            End If
            
            '2. Lieferanten abfragen
            If bLieferant_ist_enthalten = False Then
                If frmWKL30.Label1(1).Caption <> "undefiniert" Then
                    lLinr = Val(frmWKL30.Label1(1).Caption)
                    If gibtesdiesenLieferantenhier(rsrs!artnr, lLinr) Then
                        bLieferant_ist_enthalten = True
                        
                        sEtikettKürzel = ermLiefKürzelmitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettLibesnr = ermLiefLIBESNRmitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettVpe = ermLief_VPE_mitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettHS = ermLief_Handelsspanne_mitLiefvorgabe(rsrs!artnr, CStr(dVkPr), cMWST, lLinr)
                    End If
                End If
            End If
            
            '3. Lieferanten abfragen
            If bLieferant_ist_enthalten = False Then
                If frmWKL30.Label1(2).Caption <> "undefiniert" Then
                    lLinr = Val(frmWKL30.Label1(2).Caption)
                    If gibtesdiesenLieferantenhier(rsrs!artnr, lLinr) Then
                        bLieferant_ist_enthalten = True
                        
                        sEtikettKürzel = ermLiefKürzelmitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettLibesnr = ermLiefLIBESNRmitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettVpe = ermLief_VPE_mitLiefvorgabe(rsrs!artnr, lLinr)
                        sEtikettHS = ermLief_Handelsspanne_mitLiefvorgabe(rsrs!artnr, CStr(dVkPr), cMWST, lLinr)
                    End If
                End If
            End If
            
            'Lieferant gibt es bei diesem Artikel nicht dann alles nach kleinstem LEK ausrichten
            If bLieferant_ist_enthalten = False Then
                sEtikettKürzel = ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase)
                sEtikettLibesnr = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
                sEtikettVpe = ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr)
                sEtikettHS = ermLief_Handelsspanne_mitkleinstenLEKPR(rsrs!artnr, CStr(dVkPr), cMWST)
            End If
            
            rsZiel!DRUCKDATUM = Left(sEtikettKürzel, 3) & " " & cDruckdatum
            Dim cZusatz As String
            cZusatz = ermZusatztext(rsrs!artnr)
            If cZusatz = "" Then
                rsZiel!LIBESNR = sEtikettLibesnr & " VPE: " _
                & sEtikettVpe & " | " & sEtikettHS
            Else
                rsZiel!LIBESNR = sEtikettLibesnr & " VE " _
                & sEtikettVpe & " " & ermZusatztext(rsrs!artnr)
            End If
            
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
'    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
'    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DRUY set barcode13 = barcode where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25Var_Dronova"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub Update_Preis_Terminpreis(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    
    cSQL = "Update LSTEETI set vkpr = 0 where vkpr is null "
    gdBase.Execute cSQL, dbFailOnError
    
    For lcount = 0 To lAnzahl
        If CLng(acArtNr(lcount)) > 0 Then
            cSQL = "Update DRUY inner join LSTEETI on druy.artnr = LSTEETI.artnr "
            cSQL = cSQL & " set druy.VKPR = LSTEETI.vkpr, druy.VKPR_EUR = LSTEETI.vkpr where druy.artnr = " & acArtNr(lcount)
            gdBase.Execute cSQL, dbFailOnError
        End If
    
    Next lcount
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "Update_Preis_Terminpreis"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett69x38Var_Kombi(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cMWST As String

    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(50)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            cMWST = "V"
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
'            cNettoPr = nettoR(dWert, rsrs!MWST)
'            cNettoPr = Format$(cNettoPr, "#####0.00")
'            cNettoPr = Trim(cNettoPr)
'            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum
            
            Dim cZusatz As String
            cZusatz = ermZusatztext(rsrs!artnr)
            If cZusatz = "" Then
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VPE: " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " | " & ermLief_Handelsspanne_mitkleinstenLEKPR(rsrs!artnr, CStr(dVkPr), cMWST)
            Else
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VE " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " " & ermZusatztext(rsrs!artnr)
            End If
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett69x38Var_Kombi"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett81x38Var_Kombi(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cMWST As String

    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(50)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            cMWST = "V"
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
'            cNettoPr = nettoR(dWert, rsrs!MWST)
'            cNettoPr = Format$(cNettoPr, "#####0.00")
'            cNettoPr = Trim(cNettoPr)
'            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum
            
            Dim cZusatz As String
            cZusatz = ermZusatztext(rsrs!artnr)
            If cZusatz = "" Then
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VPE: " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " | " & ermLief_Handelsspanne_mitkleinstenLEKPR(rsrs!artnr, CStr(dVkPr), cMWST)
            Else
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VE " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " " & ermZusatztext(rsrs!artnr)
            End If
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett81x38Var_Kombi"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett69x38Var4(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cMWST As String

    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(50)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            cMWST = "V"
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
'            cNettoPr = nettoR(dWert, rsrs!MWST)
'            cNettoPr = Format$(cNettoPr, "#####0.00")
'            cNettoPr = Trim(cNettoPr)
'            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum
            
            Dim cZusatz As String
            cZusatz = ermZusatztext(rsrs!artnr)
            If cZusatz = "" Then
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VPE: " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " | " & ermLief_Handelsspanne_mitkleinstenLEKPR(rsrs!artnr, CStr(dVkPr), cMWST)
            Else
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase) & " VE " _
                & ermLief_VPE_mitkleinstenLEKPR(rsrs!artnr) & " " & ermZusatztext(rsrs!artnr)
            End If
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett69x38Var4"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett50x40Variante1(acArtNr() As String, lAnzahl As Long, iTage As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(20)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum & " " & ermletztVKdurch2(rsrs!artnr, iTage) & ""
            rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett50x40Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function ermletztVKdurch2(sArt As String, iTage As Integer) As Double
On Error GoTo LOKAL_ERROR

    ermletztVKdurch2 = 0#
    
    Dim sSQL As String
    Dim rsrs As Recordset

    sSQL = "Select sum(Menge) as maxi from Kassjour where artnr = " & sArt
    sSQL = sSQL & " and adate >= " & CLng(DateValue(Now) - iTage)

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            ermletztVKdurch2 = CSng(rsrs!maxi)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermletztVKdurch2"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermletztVKdurch2_APP(sArt As String, iTage As Integer, sTab As String, db As Database) As Double
On Error GoTo LOKAL_ERROR

    ermletztVKdurch2_APP = 0#
    
    Dim sSQL As String
    Dim rsrs As Recordset

    sSQL = "Select sum(Menge) as maxi from " & sTab & " where artnr = " & sArt
    

    Set rsrs = db.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            ermletztVKdurch2_APP = CSng(rsrs!maxi)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermletztVKdurch2_APP"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermletztVKdurch2_APP_SPEZIAL(sArt As String, iTage As Integer, sTab As String, db As Database) As Double
On Error GoTo LOKAL_ERROR

    ermletztVKdurch2_APP_SPEZIAL = 0#
    
    Dim sSQL As String
    Dim rsrs As Recordset

    sSQL = "Select sumMenge from " & sTab & " where artnr = " & sArt
    

    Set rsrs = db.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sumMenge) Then
            ermletztVKdurch2_APP_SPEZIAL = CSng(rsrs!sumMenge)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermletztVKdurch2_APP_SPEZIAL"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub DruckeTLPRegaletikett50x37Variante1(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett50x37Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett49x36Variante1(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett49x36Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett50x40Variante2(acArtNr() As String, lAnzahl As Long, iTage As Integer, Optional labelx As Label)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    
    labelx.Caption = ""
    labelx.Refresh
    
    Dim iCount As Integer
    iCount = 0


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(20)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    
    If lAnzahl > 5 Then
        loeschNEW "ARTLIEF", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
        loeschNEW "KASSTAGE", gdBase
        
        cSQL = "Select * into KASSTAGE from Kassjour where "
        cSQL = cSQL & " adate >= " & CLng(DateValue(Now) - iTage)
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index ARTNR on KASSTAGE (ARTNR)"
        gdBase.Execute cSQL, dbFailOnError
    
        loeschNEW "KASSTAGE", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSTAGE"
    
    
    
        loeschNEW "LISRT", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
        
        cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index LEKPR on ARTLIEF (LEKPR)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index RKZ on ARTLIEF (RKZ)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index linr on LISRT (linr)"
        gdApp.Execute cSQL, dbFailOnError
        
      
        
        
        
    End If
    
    iCount = lAnzahl
    For lcount = 0 To lAnzahl
    
        iCount = iCount - 1
        labelx.Caption = iCount
        labelx.Refresh
        
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
        
            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            
            If lAnzahl > 5 Then
                rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdApp), 3) & " " & cDruckdatum & " " & ermletztVKdurch2_APP(rsrs!artnr, iTage, "KassTage", gdApp) & ""
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdApp)
            Else
                rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum & " " & ermletztVKdurch2(rsrs!artnr, iTage) & ""
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
            End If
            
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett50x40Variante2"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeTLPRegaletikett50x40Variante3(acArtNr() As String, acSpezPreis() As String, lAnzahl As Long, iTage As Integer, Optional labelx As Label)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEAN13 As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cSpezpreis As String
    
    labelx.Caption = ""
    labelx.Refresh
    
    Dim iCount As Integer
    iCount = 0


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(20)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    
    If lAnzahl > 5 Then
        loeschNEW "ARTLIEF", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
    
        loeschNEW "KASSTAGE", gdBase
        
        cSQL = "Select * into KASSTAGE from Kassjour where "
        cSQL = cSQL & " adate >= " & CLng(DateValue(Now) - iTage)
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index ARTNR on KASSTAGE (ARTNR)"
        gdBase.Execute cSQL, dbFailOnError
    
        loeschNEW "KASSTAGE", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSTAGE"
    
    
    
        loeschNEW "LISRT", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
        
        cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index LEKPR on ARTLIEF (LEKPR)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index RKZ on ARTLIEF (RKZ)"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Create Index linr on LISRT (linr)"
        gdApp.Execute cSQL, dbFailOnError
        
      
        
        
        
    End If
    
    iCount = lAnzahl
    For lcount = 0 To lAnzahl
    
        
        
        
        iCount = iCount - 1
        labelx.Caption = iCount
        labelx.Refresh
        
        cSpezpreis = acSpezPreis(lcount)
        
        
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)


        
        If Not rsrs.EOF Then
        
            

            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!BARCODE13 = ""
            
            If Len(cEAN13) = 13 Then
                cEANCode = ean13(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If
            
            If Len(cEAN13) = 8 Then
                cEANCode = fnCodiereEANCode(cEAN13)
                rsZiel!BARCODE13 = cEANCode
            End If

            rsZiel!vkpr_EUR = rsrs!KVKPR1
            rsZiel!vkpr = cSpezpreis
            
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If
            
            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
'            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            
            If lAnzahl > 5 Then
                rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdApp), 3) & " " & cDruckdatum & " " & ermletztVKdurch2_APP(rsrs!artnr, iTage, "KassTage", gdApp) & ""
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdApp)
            Else
                rsZiel!DRUCKDATUM = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3) & " " & cDruckdatum & " " & ermletztVKdurch2(rsrs!artnr, iTage) & ""
                rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(rsrs!artnr, gdBase)
            End If
            
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett50x40Variante3"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Function Azalea_ITF_14(ByVal ITF_14 As String) As String
' I2of5Tools 25mar09 jwhiting
' Copyright 2009 Azalea Software, Inc. All rights reserved. www.azalea.com

' Creating an ITF-14 barcode in Excel
' Your input, ITF_14, is a 13-digit numeric string to be encoded as an ITF-14 symbol.

  Dim temp As String                 ' a temporary placeholder
  Dim temp2 As String                ' a temporary placeholder
  Dim chunk As String                ' loop chunk
  Dim checkDigitSubtotal As Integer  ' the check digit subtotal
  Dim checkDigit As String
  Dim i As Integer
  
  ' do the check digit calculation over all 13 digits
  checkDigitSubtotal = 3 * (Val(Mid(ITF_14, 1, 1)) + Val(Mid(ITF_14, 3, 1)) + Val(Mid(ITF_14, 5, 1)) + Val(Mid(ITF_14, 7, 1)) + Val(Mid(ITF_14, 9, 1)) + Val(Mid(ITF_14, 11, 1)) + Val(Mid(ITF_14, 13, 1)))
  checkDigitSubtotal = checkDigitSubtotal + Val(Mid(ITF_14, 2, 1)) + Val(Mid(ITF_14, 4, 1)) + Val(Mid(ITF_14, 6, 1)) + Val(Mid(ITF_14, 8, 1)) + Val(Mid(ITF_14, 10, 1)) + Val(Mid(ITF_14, 12, 1))
  checkDigit = Right(Str(300 - checkDigitSubtotal), 1)

  ITF_14 = ITF_14 & checkDigit ' now ITF_14 is a 14 (even) digit number

  temp2 = ITF_14                      ' divide input into pairs of digits
  For i = 1 To Len(ITF_14) / 2
    chunk = Left(temp2, 2)           ' grab 2 characters
    If Val(chunk) < 90 Then          ' offset into fonts' character set
      temp = temp + Chr(Val(chunk) + 33)
    ElseIf Val(chunk) = 90 Then
       temp = temp + Chr(182)
    ElseIf Val(chunk) = 91 Then
       temp = temp + Chr(183)
    ElseIf Val(chunk) > 91 Then
       temp = temp + Chr(Val(chunk) + 104)
    End If
    temp2 = Right(temp2, Len(temp2) - 2)  ' move to the next two characters
  Next i

  ' Add the start and stop bars (ASCII 171 & ASCII 172).
  Azalea_ITF_14 = Chr(171) + temp + Chr(172)

  ' Excel: B1=Azalea_ITF_14(A1)
  ' Or put another way, yourContainer.text=Azalea_ITF_14(yourInputString)
  
End Function
Public Function Interleaved_2_of_5(sZahl As String) As String
    Dim intIndex As Integer
    Dim lngSum As Long
    Dim blnSwitch As Boolean
    If IsNumeric(sZahl) Then
        For intIndex = Len(sZahl) To 1 Step -1
            If Not blnSwitch Then
                lngSum = lngSum + CInt(Mid$(sZahl, intIndex, 1)) * 3
            Else
                lngSum = lngSum + CInt(Mid$(sZahl, intIndex, 1))
            End If
            blnSwitch = Not blnSwitch
        Next
        lngSum = lngSum Mod 10
        If CBool(Len(sZahl) + Len(CStr(lngSum)) Mod 2) Then
            Interleaved_2_of_5 = CDbl(sZahl & "0" & CStr(lngSum))
        Else
            Interleaved_2_of_5 = CDbl(sZahl & CStr(lngSum))
        End If
    ElseIf Not IsEmpty(sZahl) Then
        Interleaved_2_of_5 = "Keine Zahl !"
    End If
End Function
Public Sub DruckeTLPRegaletikett40x25VarianteEdeka(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount          As Long
    Dim cSQL            As String
    Dim cSQL1           As String
    Dim rsrs            As Recordset
    Dim rsArtl          As Recordset
    Dim rsZiel          As Recordset
    Dim dWert           As Double
    Dim cEAN            As String
    Dim cEAN13          As String
    Dim cEANCode        As String
    Dim cDruckdatum     As String
    Dim cBezeich        As String
    Dim cNettoPr        As String
    Dim lAnz            As Long
    Dim dInhalt         As Double
    Dim cInhaltBez      As String
    Dim dVkPr           As Double
    Dim dGrundPreisDM   As Double
    Dim dGrundPreisEur  As Double
    Dim cGrundInhalt    As String
    Dim cLiBesNr        As String
    Dim bEdeka          As Boolean

    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", BARCODE13 Text(18)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ", LIEFKUERZEL Text(3)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            rsZiel.AddNew
            
            rsZiel!FirmaName = ""
            If gbEtiExArtikel = True Then
                rsZiel!FirmaName = "EX"
            Else
                If ermEX_INFO(rsrs!artnr) = "J" Then
                    rsZiel!FirmaName = "EX"
                End If
            End If
            
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            
            'ab hier wird nach dem richtigen EAN gesucht
            
            cEAN13 = ""
            
            If Not IsNull(rsrs!EAN) Then
                cEAN13 = rsrs!EAN 'cEAN13
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    cEAN13 = rsrs!EAN2 'cEAN13
                End If
            End If
            
            If cEAN13 = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    cEAN13 = rsrs!EAN3 'cEAN13
                End If
            End If
            
            cEAN13 = Trim(cEAN13)
            
            rsZiel!EAN1 = cEAN13
            
            cLiBesNr = ""
            bEdeka = False
            cSQL1 = "Select * from artlief where artnr = " & acArtNr(lcount) & " and linr = " & gsEdeka
            Set rsArtl = gdBase.OpenRecordset(cSQL1)
            If Not rsArtl.EOF Then
                If Not IsNull(rsArtl!LIBESNR) Then
                    cLiBesNr = rsArtl!LIBESNR
                    bEdeka = True
                End If
            End If
            rsArtl.Close
            
            rsZiel!LIBESNR = cLiBesNr
            
'            cEancode = Azalea_ITF_14(cLiBesNr)
            cEANCode = Interleaved_2_of_5(cLiBesNr)
            
            
            While Len(Trim(cLiBesNr)) < 12
                cLiBesNr = "0" & cLiBesNr
            Wend
'
'            cLiBesNr = "4" & Left(cLiBesNr, 11)
            
'            cEancode = fnMoveGutschnr2EAN13(Left(cLiBesNr, Len(cLiBesNr) - 1))
            
'            cEancode = fnCodiereEAN13Code(cEancode)
            rsZiel!BARCODE13 = cEANCode
'            rsZiel!Barcode13 = cLiBesNr
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            
            If bEdeka Then
                rsZiel!liefkuerzel = "E"
            Else
                rsZiel!liefkuerzel = Left(ermLiefKürzelmitkleinstenLEKPR(rsrs!artnr, gdBase), 3)
            End If
            
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    
    
'    cSQL = "Update DRUY set barcode13 = barcode, ean1 = 'Art# ' & artnr where barcode13 = '' "
'    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeTLPRegaletikett40x25VarianteEdeka"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Public Function ermEX_INFO(sArt As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermEX_INFO = "J"

    cSQL = "Select RKZ from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and RKZ = 'N' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        ermEX_INFO = "N"
    End If
    rsrs.Close: Set rsrs = Nothing
   
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermEX_INFO"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLiefLinrmitkleinstenLEKPR(sArt As String, db As Database) As Long
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sLinr       As String
    
    ermLiefLinrmitkleinstenLEKPR = 0
    
    sLinr = ""
    cSQL = "Select min(LEKPR) as ek , linr from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and LEKPR > 0  "
    cSQL = cSQL & " and RKZ <> 'J' "
    cSQL = cSQL & " group by linr order by min(lekpr) asc "
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!linr) Then
           sLinr = rsrs!linr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Val(sLinr) > 0 Then
        ermLiefLinrmitkleinstenLEKPR = Val(sLinr)
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLiefLinrmitkleinstenLEKPR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLiefKürzelmitkleinstenLEKPR(sArt As String, db As Database) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sLinr       As String
    
    ermLiefKürzelmitkleinstenLEKPR = ""
    
    sLinr = ""
    cSQL = "Select min(LEKPR) as ek , linr from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and LEKPR > 0  "
    cSQL = cSQL & " and RKZ <> 'J' "
    cSQL = cSQL & " group by linr order by min(lekpr) asc "
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!linr) Then
           sLinr = rsrs!linr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Val(sLinr) > 0 Then
        cSQL = "Select KUERZEL from LISRT where LINR = " & sLinr & " "
        Set rsrs = db.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Kuerzel) Then
                ermLiefKürzelmitkleinstenLEKPR = rsrs!Kuerzel
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLiefKürzelmitkleinstenLEKPR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLiefLIBESNRmitkleinstenLEKPR(sArt As String, db As Database) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLiefLIBESNRmitkleinstenLEKPR = ""
    
    cSQL = "Select min(LEKPR) as ek , libesnr from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and LEKPR > 0 "
    cSQL = cSQL & " and RKZ <> 'J' "
    cSQL = cSQL & " group by libesnr order by min(lekpr) asc "
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LIBESNR) Then
           ermLiefLIBESNRmitkleinstenLEKPR = rsrs!LIBESNR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLiefLIBESNRmitkleinstenLEKPR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLief_VPE_mitkleinstenLEKPR(sArt As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLief_VPE_mitkleinstenLEKPR = ""
    
    cSQL = "Select min(LEKPR) as ek , minmen from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and RKZ <> 'J' "
    cSQL = cSQL & " and LEKPR > 0 group by minmen order by min(lekpr) asc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!MINMEN) Then
           ermLief_VPE_mitkleinstenLEKPR = rsrs!MINMEN
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLief_VPE_mitkleinstenLEKPR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLief_Handelsspanne_mitkleinstenLEKPR(sArt As String, sKVK As String, cMWST As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim sEKpr           As String
    Dim sNettospanne    As String
    Dim dNettospanne    As Double
    
    ermLief_Handelsspanne_mitkleinstenLEKPR = "0"
    
    
    
    cSQL = "Select min(LEKPR) as ek  from ARTLIEF where ARTNR = " & sArt & " "
    cSQL = cSQL & " and LEKPR > 0 "
    cSQL = cSQL & " and RKZ <> 'J' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ek) Then
           sEKpr = rsrs!ek
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    sNettospanne = NettospanneInProzent(sKVK, sEKpr, cMWST)
    dNettospanne = CDbl(sNettospanne)
        
    ermLief_Handelsspanne_mitkleinstenLEKPR = Fix(dNettospanne)
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLief_Handelsspanne_mitkleinstenLEKPR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function gibtesdiesenLieferantenhier(sArt As String, lLinr As Long) As Boolean
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    gibtesdiesenLieferantenhier = False
    
    cSQL = "Select * from ARTLIEF where ARTNR = " & sArt & " and Linr = " & lLinr & " and RKZ <> 'J' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LIBESNR) Then
           gibtesdiesenLieferantenhier = True
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "gibtesdiesenLieferantenhier"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function



Public Function ermLiefKürzelmitLiefvorgabe(sArt As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim bgefunden   As Boolean
    
    ermLiefKürzelmitLiefvorgabe = ""
    bgefunden = False
    cSQL = "Select * from ARTLIEF where ARTNR = " & sArt & " and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        bgefunden = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bgefunden Then
        cSQL = "Select KUERZEL from LISRT where LINR = " & lLinr & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Kuerzel) Then
                ermLiefKürzelmitLiefvorgabe = rsrs!Kuerzel
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLiefKürzelmitLiefvorgabe"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLiefLIBESNRmitLiefvorgabe(sArt As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLiefLIBESNRmitLiefvorgabe = ""
    
    cSQL = "Select libesnr from ARTLIEF where ARTNR = " & sArt & " and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LIBESNR) Then
           ermLiefLIBESNRmitLiefvorgabe = rsrs!LIBESNR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLiefLIBESNRmitLiefvorgabe"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLief_VPE_mitLiefvorgabe(sArt As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    ermLief_VPE_mitLiefvorgabe = ""
    
    cSQL = "Select minmen from ARTLIEF where ARTNR = " & sArt & " and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!MINMEN) Then
           ermLief_VPE_mitLiefvorgabe = rsrs!MINMEN
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLief_VPE_mitLiefvorgabe"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLief_Handelsspanne_mitLiefvorgabe(sArt As String, sKVK As String, cMWST As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim sEKpr           As String
    Dim sNettospanne    As String
    Dim dNettospanne    As Double
    
    ermLief_Handelsspanne_mitLiefvorgabe = "0"
    
    cSQL = "Select LEKPR as ek  from ARTLIEF where ARTNR = " & sArt & " and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ek) Then
           sEKpr = rsrs!ek
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    sNettospanne = NettospanneInProzent(sKVK, sEKpr, cMWST)
    dNettospanne = CDbl(sNettospanne)
        
    ermLief_Handelsspanne_mitLiefvorgabe = Fix(dNettospanne)
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "ermLief_Handelsspanne_mitLiefvorgabe"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub spezial_DruckeTLPRegaletikett70x35Variante1(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim cNettoPr As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim dLVK As Double
    Dim cLVK As String
    Dim cArtNr As String


    loeschNEW "DRUY", gdBase
    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then

        
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            rsZiel!EAN1 = rsrs!EAN 'cEAN

            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode

            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            cNettoPr = nettoR(dWert, rsrs!MWST)
            cNettoPr = Format$(cNettoPr, "#####0.00")
            cNettoPr = Trim(cNettoPr)
            rsZiel!NEPR = cNettoPr

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            
            rsZiel.Update
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    SpalteAnfuegenNEW "DRUY", "LVK", "Text(20)", gdBase
    SpalteAnfuegenNEW "DRUY", "NETTOPR", "double", gdBase
            
    cSQL = "Select * from DRUY "
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    If Not rsZiel.EOF Then
        rsZiel.MoveFirst
        Do While Not rsZiel.EOF
    
            If Not IsNull(rsZiel!artnr) Then
                cArtNr = rsZiel!artnr
            End If
            
            If Val(cArtNr) > 0 Then
        
                cSQL = "Select * from Artikel where artnr = " & cArtNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                
                    If Not IsNull(rsrs!KVKPR1) Then
                        dWert = rsrs!KVKPR1
                    End If
                
                    cNettoPr = nettoR(dWert, rsrs!MWST)
                    cNettoPr = Format$(cNettoPr, "#####0.00")
                    cNettoPr = Trim(cNettoPr)
                    
                    dLVK = Format$(CDbl(cNettoPr) * 80 / 100, "#####0.00")
                    
                    cLVK = SwapStr(dLVK, ",", "")
                    cLVK = "HWHP0000" & cLVK

                    rsZiel.Edit
                    rsZiel!NETTOPR = cNettoPr
                    rsZiel!LVK = cLVK
                    rsZiel.Update
                    
                End If
                rsrs.Close: Set rsrs = Nothing
            
            End If
            rsZiel.MoveNext
        Loop
    End If
    rsZiel.Close: Set rsZiel = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "spezial_DruckeTLPRegaletikett70x35Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeGrundPreisEtikettenWKL30Jebe(acArtNr() As String, lAnzahl As Long, sArt As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cKVkPr1 As String
    Dim cNettoPr As String

    Dim iFileNr As Integer
    
    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    Select Case sArt
        Case Is = "BRUTTO"
            cSQL = "Delete from DRUBRU "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Select * from DRUBRU"
        Case Is = "NETTO"
            cSQL = "Delete from DRUNET "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Select * from DRUNET"
    End Select
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            rsZiel!EAN1 = rsrs!EAN 'cEAN

            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode

            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            
            
            If sArt = "NETTO" Then
                cKVkPr1 = Format$(dWert, "#####0.00")
                cKVkPr1 = SwapStr(cKVkPr1, ",", ".")
                cKVkPr1 = Trim(cKVkPr1)
            
                cNettoPr = nettoR(dWert, rsrs!MWST)
                cNettoPr = Format$(cNettoPr, "#####0.00")
                cNettoPr = Trim(cNettoPr)
                rsZiel!NEPR = cNettoPr
            End If
            
            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If

            '//DRU_GRUN
            rsZiel.Update
            
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
   
   'Achtung hier Änderung Hintergrundtabellen = DRUBRU und DRUNET
    Select Case sArt
        Case Is = "BRUTTO"
        
            If Modul6.FindFile(gcDBPfad, "aWKL30as.rpt") Then
                reportbildschirmToPrinterETI "aWKL30as", gcEtikettenDrucker, True
            Else
                reportbildschirm "WKL017", "aWKL30a"
            End If
            
        Case Is = "NETTO"
            If Modul6.FindFile(gcDBPfad, "aWKL30bs.rpt") Then
                reportbildschirmToPrinterETI "aWKL30bs", gcEtikettenDrucker, True
    
            Else
                reportbildschirm "WKL017", "aWKL30b"
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "DruckeGrundPreisEtikettenWKL30Jebe"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

        Fehlermeldung1
    End If
End Sub
Public Sub DruckeNettoStrichcode(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cKVkPr1 As String
    Dim cNettoPr As String

    Dim iFileNr As Integer

    loeschNEW "DRU_GRUN", gdBase

    cSQL = "Create Table DRU_GRUN ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "

    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRU_GRUN"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
                
                
                
'                If sart = "NETTO" Then
                    cKVkPr1 = Format$(dWert, "#####0.00")
                    cKVkPr1 = SwapStr(cKVkPr1, ",", ".")
                    cKVkPr1 = Trim(cKVkPr1)
                
                    cNettoPr = nettoR(dWert, rsrs!MWST)
                    cNettoPr = Format$(cNettoPr, "#####0.00")
    
                    cNettoPr = Trim(cNettoPr)
                    rsZiel!NEPR = cNettoPr
'                End If
                
                
                
                
                
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                '//DRU_GRUN
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    If Modul6.FindFile(gcDBPfad, "aWKL30bs.rpt") Then
        reportbildschirmToPrinterETI "aWKL30bs", gcEtikettenDrucker, True

    Else
        reportbildschirm "WKL017", "aWKL30b"
    End If
            

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "DruckeNettoStrichcode"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

        Fehlermeldung1
    End If
End Sub
Public Sub DruckeStrichcodeY(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    Dim cKVkPr1 As String
    Dim cNettoPr As String

    Dim iFileNr As Integer

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "

    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                '//DRU_GRUN
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul4"
        Fehler.gsFunktion = "DruckeStrichcodeY"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

        Fehlermeldung1
    End If
End Sub
Public Sub DruckeLFNREtikett()
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cBezeich As String
    Dim lMaxlfnr As Long
    Dim cDatum As String

    Dim iFileNr As Integer

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(10)"
    cSQL = cSQL & ") "

    gdBase.Execute cSQL, dbFailOnError

    

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    cSQL1 = "Select max(lfnr)as maxi from zugang "
    Set rsrs = gdBase.OpenRecordset(cSQL1)
        lMaxlfnr = rsrs!maxi
    rsrs.Close: Set rsrs = Nothing
    
    cSQL1 = "Select * from zugang where lfnr = " & lMaxlfnr
    Set rsrs = gdBase.OpenRecordset(cSQL1)
    rsZiel.AddNew
    rsZiel!artnr = rsrs!artnr
    If Not IsNull(rsrs!BEZEICH) Then
        cBezeich = rsrs!BEZEICH
    Else
        cBezeich = ""
    End If
    cBezeich = fnEntferneLeerzeichen(cBezeich)
    rsZiel!EAN = lMaxlfnr
    rsZiel!BEZEICH = cBezeich
    cDatum = Format(rsrs!ADATE, "dd.mm.yy")
    cDatum = SwapStr(cDatum, ".", "")
    cDatum = Right(cDatum, 2) & Left(cDatum, 2) & Mid(cDatum, 3, 2)
    rsZiel!DRUCKDATUM = cDatum
    rsZiel!EAN1 = rsrs!artnr
    rsZiel!Barcode = rsrs!linr
    rsZiel.Update
               
    rsrs.Close: Set rsrs = Nothing
    
    rsZiel.Close: Set rsZiel = Nothing

    reportbildschirmToPrinterETI "alfnr", gcEtikettenDrucker, False
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeLFNREtikett"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Public Function nettoR(dWert As Double, sMWST As String) As Double
On Error GoTo LOKAL_ERROR

Select Case sMWST
    Case "V"
        nettoR = (dWert * 100) / (100 + gdMWStV)
    Case "E"
        nettoR = (dWert * 100) / (100 + gdMWStE)
    Case "O"
        nettoR = (dWert * 100) / (100 + gdMWStO)
End Select

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "nettoR"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
  Public Function SyncShell(CommandLine As String, TimeOut As Long, WaitForInputIdle As Boolean, Hide As Boolean) As Boolean
      Dim hProcess As Long
      Dim ret As Long
      Dim nMilliseconds As Long
      If TimeOut > 0 Then
          nMilliseconds = TimeOut
      Else
          nMilliseconds = INFINITE
      End If
      hProcess = StartProcess(CommandLine, Hide)
      If WaitForInputIdle Then
          ' Warten, bis die eingeschlossene Anwendung
          ' mit dem Erstellen ihrer Schnittstelle fertig ist:
          ret = InputIdle(hProcess, nMilliseconds)
      Else
          ' Warten, bis die eingeschlossene Anwendung beendet ist:
          ret = WaitForSingleObject(hProcess, nMilliseconds)
      End If
      CloseHandle hProcess
      ' "True" zurückgeben, wenn die Anwendung fertig ist.
      ' Andernfalls Zeitüberschreitung oder Fehler.
      SyncShell = (ret = WAIT_OBJECT_0)
  End Function
  
  Public Function StartProcess(CommandLine As String, Hide As Boolean) As Long
      Const STARTF_USESHOWWINDOW As Long = &H1
      Const SW_HIDE As Long = 0
      
      Dim proc As PROCESS_INFORMATION
      Dim Start As STARTUPINFO
      ' STARTUPINFO-Struktur initialisieren:
      Start.cb = Len(Start)
      If Hide Then
          Start.dwFlags = STARTF_USESHOWWINDOW
          Start.wShowWindow = SW_HIDE
      End If
      ' Eingeschlossene Anwendung starten:
      CreateProcessA 0&, CommandLine, 0&, 0&, 1&, _
          NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc
      StartProcess = proc.hProcess
  End Function
Public Function fnFileSize(cdatei As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer

    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    fnFileSize = LOF(iFileNr)
    Close iFileNr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "fnFileSize"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub DruckeSchmucketikett69x14Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                rsZiel!vkpr_EUR = rsrs!vkpr
'                If Not IsNull(rsrs!KVKPR1) Then
'                    dVkPr = rsrs!KVKPR1
'                Else
'                    dVkPr = 0
'                End If
'
'                If dVkPr <> 0 Then
'                    dWert = dVkPr
'                End If
'                rsZiel!vkpr_eur = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeSchmucketikett69x14Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett40x18Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett40x18Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett40x18Variante5(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
                
                rsZiel!vkpr_EUR = rsrs!vkpr
    
'                If dVkPr <> 0 Then
'                    dWert = dVkPr
'                End If
'                rsZiel!vkpr_eur = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett40x18Variante5"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett44x21Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett44x21Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett48x18Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett48x18Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett45x23Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
                rsZiel!vkpr_EUR = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett45x23Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett51x19Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                rsZiel!vkpr_EUR = rsrs!vkpr
                
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
'                rsZiel!vkpr_eur = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
                
                
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett51x19Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett49x19Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                rsZiel!vkpr_EUR = rsrs!vkpr
                
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
'                rsZiel!vkpr_eur = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
                
                
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett49x19Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEtikett35x15Variante1(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String

    loeschNEW "DRUY", gdBase

    cSQL = "Create Table DRUY ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

    cSQL = "Select * from DRUY"
    Set rsZiel = gdBase.OpenRecordset(cSQL)

    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If Not rsrs.EOF Then
            For lAnz = 1 To Val(acAnzEti(lcount))
        
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
    
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
                
                rsZiel!EAN = cEAN
                rsZiel!EAN1 = rsrs!EAN 'cEAN
    
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
    
                rsZiel!LIBESNR = rsrs!LIBESNR
    
                rsZiel!vkpr = rsrs!KVKPR1
                rsZiel!vkpr_EUR = rsrs!vkpr
                
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
    
                If dVkPr <> 0 Then
                    dWert = dVkPr
                End If
'                rsZiel!vkpr_eur = dWert
    
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
                
                
    
                rsZiel.Update
               
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul4"
    Fehler.gsFunktion = "DruckeEtikett35x15Variante1"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

