Attribute VB_Name = "mdl_ZVT2"
Option Explicit



Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Eingabeparameter
Global COM As String
Global ComSpeed As Long
Global ComStop As Long
Global IP As String
Global Port As Long
Global KasseNr As Long
Global Kassedruck As Long
Global Protokollpfad As String
Global Betrag As Long
Global Lizenz As String
Global Passwort As String
Global Provider As Long
Global Funktion As Long
Global Test As Long
Global dialog As Long

Global Kartennummer As String
Global Kartegueltig As String
Global KarteCVC As String

Global StornoBetrag As Long
Global StornoBelegNr As Long




' Ausgabeparameter
Global Ergebnis As Long
Global ErgebnisText As String
Global ErgebnisLang As String
Global Autorisierungsergebnis As String
Global Kundenbeleg As String
Global Haendlerbeleg As String
Global HaendlerbelegPROF As String
Global Kartentyp As Long
Global KartentypText As String
Global gsKartenNummer As String
Global gsKarteGueltig As String

' Hilfsfunktionen zum Lesen und Schreiben in der Registry

Function RegRead(RegKey As String) As Variant
    Dim myObject As Object
    Dim Path As String
    Path = "HKEY_CURRENT_USER\SOFTWARE\GUB\ZVT\"
    On Error Resume Next
    Set myObject = CreateObject("WScript.Shell")
    RegRead = myObject.RegRead(Path & RegKey)
End Function

Sub RegWriteSZ(RegKey As String, value As String)
    Dim myObject As Object
    Dim Path As String
    Path = "HKEY_CURRENT_USER\SOFTWARE\GUB\ZVT\"
    On Error Resume Next
    Set myObject = CreateObject("WScript.Shell")
    myObject.RegWrite Path & RegKey, value, "REG_SZ"
End Sub

Sub RegWriteDWORD(RegKey As String, value As Long)
    Dim myObject As Object
    Dim Path As String
    Path = "HKEY_CURRENT_USER\SOFTWARE\GUB\ZVT\"
    On Error Resume Next
    Set myObject = CreateObject("WScript.Shell")
    myObject.RegWrite Path & RegKey, value, "REG_DWORD"
End Sub
Sub Zahlen()

    Dim Startpfad As String
    Dim Aktiv As Long

    RegWriteDWORD "Funktion", Funktion ' 0 = Zahlung, 1 = Diagnose, 2 = Kassenschnitt,3 = Storno
    RegWriteSZ "COM", COM
    RegWriteDWORD "ComSpeed", ComSpeed
    RegWriteDWORD "ComStop", ComStop
    RegWriteSZ "IP", IP
    RegWriteDWORD "Port", Port
    RegWriteDWORD "Betrag", Betrag
    RegWriteSZ "Protokoll", Protokollpfad
    RegWriteDWORD "Test", Test
    RegWriteDWORD "KasseNr", KasseNr
    RegWriteDWORD "Kassedruck", Kassedruck
    RegWriteSZ "Lizenz", Lizenz
    RegWriteSZ "Passwort", Passwort
    RegWriteDWORD "Provider", Provider
    RegWriteDWORD "Dialog", dialog
    
    RegWriteSZ "Kartennummer", Kartennummer
    RegWriteSZ "Kartegueltig", Kartegueltig
    RegWriteSZ "KarteCVC", KarteCVC

    Startpfad = RegRead("Start") ' Programmpfad dynamisch auslesen

    'EasyZVT starten
    Shell Startpfad
    
    PauseSi 0.5

    'auf Programmende aktiv warten - in Access ist kein warten auf Ende des Shellaufrufs möglich
    Do
        Sleep 1000 ' 100 ms warten
        Aktiv = RegRead("Aktiv")
        DoEvents
    Loop While Aktiv = 1
    
    giZVT2_Fehler = RegRead("Ergebnis")
    
    
   
    ' Ergebnisse auslesen
    Ergebnis = RegRead("Ergebnis")
    ErgebnisText = RegRead("ErgebnisText")
    ErgebnisLang = RegRead("ErgebnisLang")
    Autorisierungsergebnis = RegRead("Autorisierungsergebnis")
    Kundenbeleg = RegRead("Drucktext")
    
    
    
    If gbZVT2_KBDrucken = True Then
        Haendlerbeleg = RegRead("Drucktext2")
        HaendlerbelegPROF = RegRead("Haendlerbeleg")
    Else
        Haendlerbeleg = RegRead("Drucktext2")
    End If
    
    
    
    
    
    Kartentyp = RegRead("Kartentyp")
    KartentypText = RegRead("KartentypLang")
    
    gsKartenNummer = RegRead("Kartennummer")
    gsKarteGueltig = RegRead("Kartegueltig")
    


    If iWelchekarte = 1 Or iWelchekarte = 0 Then
    
            If gbZVT2_Kartenwahl = False Then
                Select Case Kartentyp
                   
                   Case 2 'EC
                        gcKreditKarte = "EC"
                        
                    Case 5 'girocard
                        gcKreditKarte = "GI"
                        
                    Case 6 'MasterCard
                        gcKreditKarte = "MC"
                        
                    Case 8 'American Express
                        gcKreditKarte = "AE"
                        
                    Case 10 'Visa
                        gcKreditKarte = "VI"
        
                    Case 12 'Diners
                        gcKreditKarte = "DI"
                        
                    Case Else
                        gcKreditKarte = "SO"
                    
                End Select
            End If
    End If
    
    
    
    If iWelchekarte = 2 Then
    
            If gbZVT2_Kartenwahl = False Then
                Select Case Kartentyp
                   
                   Case 2 'EC
                        gcKreditKarte2 = "EC"
                        
                    Case 5 'girocard
                        gcKreditKarte2 = "GI"
                        
                    Case 6 'MasterCard
                        gcKreditKarte2 = "MC"
                        
                    Case 8 'American Express
                        gcKreditKarte2 = "AE"
                        
                    Case 10 'Visa
                        gcKreditKarte2 = "VI"
        
                    Case 12 'Diners
                        gcKreditKarte2 = "DI"
                        
                    Case Else
                        gcKreditKarte2 = "SO"
                    
                End Select
            End If
    End If
    
    
    
    
    schreibeProtokollUNITXT Kartentyp & ";" & KartentypText & ";" & gsKartenNummer & ";" & gsKarteGueltig, "ZVT_KARTEN"
    
    schreibeProtokollUNITXT Kundenbeleg, "ZVT_KUNDENBELEG"
    
    If gbZVT2_KBDrucken = True Then
        schreibeProtokollUNITXT HaendlerbelegPROF, "ZVT_HAENDLERBELEG"
    Else
        schreibeProtokollUNITXT Haendlerbeleg, "ZVT_HAENDLERBELEG"
    End If
    
    
    
    

    
    If gbZVT2_KBDrucken = True Then
    
        Dim cFeld As String
        Dim sBelegtext As String
        Dim iDruckzeilen_count As Integer
        ReDim cDruckZeile(1 To 1) As String
        
        sBelegtext = Kundenbeleg
        iDruckzeilen_count = 0
        ReDim cDruckZeile(1 To 1) As String

        Dim sArray() As String
        Dim i As Integer
        sArray = Split(sBelegtext, vbLf)
        
        For i = 0 To UBound(sArray)
            iDruckzeilen_count = iDruckzeilen_count + 1
            ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
            cFeld = sArray(i)
            cDruckZeile(iDruckzeilen_count) = cFeld
        Next i

        DruckeTextArray cDruckZeile(), iDruckzeilen_count
    
    End If
    
    If gbZVT2_HBDrucken = True Then
    
        Dim cFeld_HB As String
        Dim sBelegtext_HB As String
        Dim iDruckzeilen_count_HB As Integer
        ReDim cDruckZeile_HB(1 To 1) As String
        
        sBelegtext_HB = HaendlerbelegPROF
        iDruckzeilen_count_HB = 0
        ReDim cDruckZeile_HB(1 To 1) As String

        Dim sArray_HB() As String
        Dim i_HB As Integer
        sArray_HB = Split(sBelegtext_HB, vbLf)
        
        For i_HB = 0 To UBound(sArray_HB)
            iDruckzeilen_count_HB = iDruckzeilen_count_HB + 1
            ReDim Preserve cDruckZeile_HB(1 To iDruckzeilen_count_HB) As String
            cFeld_HB = sArray_HB(i_HB)
            cDruckZeile_HB(iDruckzeilen_count_HB) = cFeld_HB
        Next i_HB

        DruckeTextArray cDruckZeile_HB(), iDruckzeilen_count_HB
    
    End If
    
    
    
End Sub
Sub Storno()

    Dim Startpfad As String
    Dim Aktiv As Long

    RegWriteDWORD "Funktion", Funktion ' 0 = Zahlung, 1 = Diagnose, 2 = Kassenschnitt,3 = Storno
    RegWriteSZ "COM", COM
    RegWriteDWORD "ComSpeed", ComSpeed
    RegWriteDWORD "ComStop", ComStop
    RegWriteSZ "IP", IP
    RegWriteDWORD "Port", Port
    
    RegWriteDWORD "StornoBetrag", StornoBetrag
    RegWriteDWORD "StornoBelegNr", StornoBelegNr
    
    RegWriteDWORD "Betrag", StornoBetrag
    
    RegWriteSZ "Protokoll", Protokollpfad
    RegWriteDWORD "Test", Test
    RegWriteDWORD "KasseNr", KasseNr
    RegWriteDWORD "Kassedruck", Kassedruck
    RegWriteSZ "Lizenz", Lizenz
    RegWriteSZ "Passwort", Passwort
    RegWriteDWORD "Provider", Provider
    RegWriteDWORD "Dialog", dialog
    
    RegWriteSZ "Kartennummer", Kartennummer
    RegWriteSZ "Kartegueltig", Kartegueltig
    RegWriteSZ "KarteCVC", KarteCVC

    Startpfad = RegRead("Start") ' Programmpfad dynamisch auslesen

    'EasyZVT starten
    Shell Startpfad
    
    PauseSi 0.5

    'auf Programmende aktiv warten - in Access ist kein warten auf Ende des Shellaufrufs möglich
    Do
        Sleep 1000 ' 100 ms warten
        Aktiv = RegRead("Aktiv")
        DoEvents
    Loop While Aktiv = 1

    giZVT2_Fehler = RegRead("Ergebnis")
    
    ' Ergebnisse auslesen
    Ergebnis = RegRead("Ergebnis")
    ErgebnisText = RegRead("ErgebnisText")
    ErgebnisLang = RegRead("ErgebnisLang")
    Autorisierungsergebnis = RegRead("Autorisierungsergebnis")
    Kundenbeleg = RegRead("Drucktext")
    If gbZVT2_KBDrucken = True Then
        Haendlerbeleg = RegRead("Drucktext2")
        HaendlerbelegPROF = RegRead("Haendlerbeleg")
    Else
        Haendlerbeleg = RegRead("Drucktext2")
    End If
    Kartentyp = RegRead("Kartentyp")
    KartentypText = RegRead("KartentypLang")
    
    
    
    
    If iWelchekarte = 1 Or iWelchekarte = 0 Then
    
            If gbZVT2_Kartenwahl = False Then
                Select Case Kartentyp
                   
                   Case 2 'EC
                        gcKreditKarte = "EC"
                        
                    Case 5 'girocard
                        gcKreditKarte = "GI"
                        
                    Case 6 'MasterCard
                        gcKreditKarte = "MC"
                        
                    Case 8 'American Express
                        gcKreditKarte = "AE"
                        
                    Case 10 'Visa
                        gcKreditKarte = "VI"
        
                    Case 12 'Diners
                        gcKreditKarte = "DI"
                        
                    Case Else
                        gcKreditKarte = "SO"
                        
                End Select
            End If
    
    End If
    
    
        If iWelchekarte = 2 Then
    
            If gbZVT2_Kartenwahl = False Then
                Select Case Kartentyp
                   
                   Case 2 'EC
                        gcKreditKarte2 = "EC"
                        
                    Case 5 'girocard
                        gcKreditKarte2 = "GI"
                        
                    Case 6 'MasterCard
                        gcKreditKarte2 = "MC"
                        
                    Case 8 'American Express
                        gcKreditKarte2 = "AE"
                        
                    Case 10 'Visa
                        gcKreditKarte2 = "VI"
        
                    Case 12 'Diners
                        gcKreditKarte2 = "DI"
                        
                    Case Else
                        gcKreditKarte2 = "SO"
                        
                End Select
            End If
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    schreibeProtokollUNITXT Kundenbeleg, "ZVT_KUNDENBELEG"
    
    If gbZVT2_KBDrucken = True Then
        schreibeProtokollUNITXT HaendlerbelegPROF, "ZVT_HAENDLERBELEG"
    Else
        schreibeProtokollUNITXT Haendlerbeleg, "ZVT_HAENDLERBELEG"
    End If
    
    If gbZVT2_KBDrucken = True Then
    
        Dim cFeld As String
        Dim sBelegtext As String
        Dim iDruckzeilen_count As Integer
        ReDim cDruckZeile(1 To 1) As String
        
        sBelegtext = Kundenbeleg
        iDruckzeilen_count = 0
        ReDim cDruckZeile(1 To 1) As String

        Dim sArray() As String
        Dim i As Integer
        sArray = Split(sBelegtext, vbLf)
        
        For i = 0 To UBound(sArray)
            iDruckzeilen_count = iDruckzeilen_count + 1
            ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
            cFeld = sArray(i)
            cDruckZeile(iDruckzeilen_count) = cFeld
        Next i

        DruckeTextArray cDruckZeile(), iDruckzeilen_count
    
    End If
    
    If gbZVT2_HBDrucken = True Then
    
        Dim cFeld_HB As String
        Dim sBelegtext_HB As String
        Dim iDruckzeilen_count_HB As Integer
        ReDim cDruckZeile_HB(1 To 1) As String
        
        sBelegtext_HB = HaendlerbelegPROF
        iDruckzeilen_count_HB = 0
        ReDim cDruckZeile_HB(1 To 1) As String

        Dim sArray_HB() As String
        Dim i_HB As Integer
        sArray_HB = Split(sBelegtext_HB, vbLf)
        
        For i_HB = 0 To UBound(sArray_HB)
            iDruckzeilen_count_HB = iDruckzeilen_count_HB + 1
            ReDim Preserve cDruckZeile_HB(1 To iDruckzeilen_count_HB) As String
            cFeld_HB = sArray_HB(i_HB)
            cDruckZeile_HB(iDruckzeilen_count_HB) = cFeld_HB
        Next i_HB

        DruckeTextArray cDruckZeile_HB(), iDruckzeilen_count_HB
    
    End If
    
    
    
    
End Sub
Sub Kassenschnitt()

    Dim Startpfad As String
    Dim Aktiv As Long

    RegWriteDWORD "Funktion", Funktion ' 0 = Zahlung, 1 = Diagnose, 2 = Kassenschnitt,3 = Storno
    RegWriteSZ "COM", COM
    RegWriteDWORD "ComSpeed", ComSpeed
    RegWriteDWORD "ComStop", ComStop
    RegWriteSZ "IP", IP
    RegWriteDWORD "Port", Port
'    RegWriteDWORD "Betrag", Betrag
    RegWriteSZ "Protokoll", Protokollpfad
    RegWriteDWORD "Test", Test
    RegWriteDWORD "KasseNr", KasseNr
    RegWriteDWORD "Kassedruck", Kassedruck
    RegWriteSZ "Lizenz", Lizenz
    RegWriteSZ "Passwort", Passwort
    RegWriteDWORD "Provider", Provider
    RegWriteDWORD "Dialog", dialog

    Startpfad = RegRead("Start") ' Programmpfad dynamisch auslesen

    'EasyZVT starten
    Shell Startpfad
    
    PauseSi 0.5

    'auf Programmende aktiv warten - in Access ist kein warten auf Ende des Shellaufrufs möglich
    Do
        Sleep 100 ' 100 ms warten
        Aktiv = RegRead("Aktiv")
        DoEvents
    Loop While Aktiv = 1
    
    giZVT2_Fehler = RegRead("Ergebnis")

    ' Ergebnisse auslesen
    Ergebnis = RegRead("Ergebnis")
    ErgebnisText = RegRead("ErgebnisText")
    ErgebnisLang = RegRead("ErgebnisLang")
    Autorisierungsergebnis = RegRead("Autorisierungsergebnis")
    Kundenbeleg = RegRead("Drucktext")
    
    If gbZVT2_KBDrucken = True Then
        Haendlerbeleg = RegRead("Drucktext2")
        HaendlerbelegPROF = RegRead("Haendlerbeleg")
    Else
        Haendlerbeleg = RegRead("Drucktext2")
    End If
    
    
    Kartentyp = RegRead("Kartentyp")
    KartentypText = RegRead("KartentypLang")
    
    
    
    If gbZVT2_KBDrucken = True Then
    
        Dim cFeld As String
        Dim sBelegtext As String
        Dim iDruckzeilen_count As Integer
        ReDim cDruckZeile(1 To 1) As String
        
        sBelegtext = Kundenbeleg
        iDruckzeilen_count = 0
        ReDim cDruckZeile(1 To 1) As String

        Dim sArray() As String
        Dim i As Integer
        sArray = Split(sBelegtext, vbLf)
        
        For i = 0 To UBound(sArray)
            iDruckzeilen_count = iDruckzeilen_count + 1
            ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
            cFeld = sArray(i)
            cDruckZeile(iDruckzeilen_count) = cFeld
        Next i

        DruckeTextArray cDruckZeile(), iDruckzeilen_count
        
        
        
        
        
        
        
        Dim cFeld1 As String
        Dim sBelegtext1 As String
        Dim iDruckzeilen_count1 As Integer
        ReDim cDruckZeile1(1 To 1) As String
        
        sBelegtext1 = Haendlerbeleg
        iDruckzeilen_count1 = 0
        ReDim cDruckZeile1(1 To 1) As String

        Dim sArray1() As String
        Dim i1 As Integer
        sArray1 = Split(sBelegtext1, vbLf)
        
        For i1 = 0 To UBound(sArray1)
            iDruckzeilen_count1 = iDruckzeilen_count1 + 1
            ReDim Preserve cDruckZeile1(1 To iDruckzeilen_count1) As String
            cFeld1 = sArray1(i1)
            cDruckZeile1(iDruckzeilen_count1) = cFeld1
        Next i1

        DruckeTextArray cDruckZeile1(), iDruckzeilen_count1
        
        
        
        Dim cFeld2 As String
        Dim sBelegtext2 As String
        Dim iDruckzeilen_count2 As Integer
        ReDim cDruckZeile2(1 To 1) As String
        
        sBelegtext2 = HaendlerbelegPROF
        iDruckzeilen_count2 = 0
        ReDim cDruckZeile2(1 To 1) As String

        Dim sArray2() As String
        Dim i2 As Integer
        sArray2 = Split(sBelegtext2, vbLf)
        
        For i2 = 0 To UBound(sArray2)
            iDruckzeilen_count2 = iDruckzeilen_count2 + 1
            ReDim Preserve cDruckZeile2(1 To iDruckzeilen_count2) As String
            cFeld2 = sArray2(i2)
            cDruckZeile2(iDruckzeilen_count2) = cFeld2
        Next i2

        DruckeTextArray cDruckZeile2(), iDruckzeilen_count2
    
    
    End If
    
    
    
    
    
    
End Sub
Sub Belegwiederholung()

    Dim Startpfad As String
    Dim Aktiv As Long

    RegWriteDWORD "Funktion", Funktion ' 0 = Zahlung, 1 = Diagnose, 2 = Kassenschnitt,3 = Storno
    RegWriteSZ "COM", COM
    RegWriteDWORD "ComSpeed", ComSpeed
    RegWriteDWORD "ComStop", ComStop
    RegWriteSZ "IP", IP
    RegWriteDWORD "Port", Port
'    RegWriteDWORD "Betrag", Betrag
    RegWriteSZ "Protokoll", Protokollpfad
    RegWriteDWORD "Test", Test
    RegWriteDWORD "KasseNr", KasseNr
    RegWriteDWORD "Kassedruck", Kassedruck
    RegWriteSZ "Lizenz", Lizenz
    RegWriteSZ "Passwort", Passwort
    RegWriteDWORD "Provider", Provider
    RegWriteDWORD "Dialog", dialog

    Startpfad = RegRead("Start") ' Programmpfad dynamisch auslesen

    'EasyZVT starten
    Shell Startpfad
    
    PauseSi 0.5

    'auf Programmende aktiv warten - in Access ist kein warten auf Ende des Shellaufrufs möglich
    Do
        Sleep 100 ' 100 ms warten
        Aktiv = RegRead("Aktiv")
        DoEvents
    Loop While Aktiv = 1
    
    giZVT2_Fehler = RegRead("Ergebnis")

    ' Ergebnisse auslesen
    Ergebnis = RegRead("Ergebnis")
    ErgebnisText = RegRead("ErgebnisText")
    ErgebnisLang = RegRead("ErgebnisLang")
    Autorisierungsergebnis = RegRead("Autorisierungsergebnis")
    Kundenbeleg = RegRead("Drucktext")
    Haendlerbeleg = RegRead("Drucktext2")
    Kartentyp = RegRead("Kartentyp")
    KartentypText = RegRead("KartentypLang")
    
    
    
    If gbZVT2_KBDrucken = True Then
    
        Dim cFeld As String
        Dim sBelegtext As String
        Dim iDruckzeilen_count As Integer
        ReDim cDruckZeile(1 To 1) As String
        
        sBelegtext = Kundenbeleg
        iDruckzeilen_count = 0
        ReDim cDruckZeile(1 To 1) As String

        Dim sArray() As String
        Dim i As Integer
        sArray = Split(sBelegtext, vbLf)
        
        For i = 0 To UBound(sArray)
            iDruckzeilen_count = iDruckzeilen_count + 1
            ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
            cFeld = sArray(i)
            cDruckZeile(iDruckzeilen_count) = cFeld
        Next i

        DruckeTextArray cDruckZeile(), iDruckzeilen_count
    
    End If
    
    
    
    
    
    
End Sub

Public Function Zahlung_ZVT2(sBetrag As String, bmitMess As Boolean, _
Optional sKartennummer As String = "", Optional sKartegueltig As String = "", Optional sKarteCVC As String = "") As String
    On Error GoTo LOKAL_ERROR

    Zahlung_ZVT2 = ""
    
    
    'Beispielwerte setzen
    COM = "LAN" ' Alternativ "COM" (automatische Com-Port-Erkennung) oder z.B. COM11 (fixer COM-Port)
    ComSpeed = 9600 ' Geräteabhängig, Standard = 9600
    ComStop = 2 ' Geräteabhängig 1 oder 2, Standard = 2
    IP = gZVT2_IP ' wenn IP verwendet wird, dann bitte IP-Adresse am EC-Gerät fest einstellen, Standard ist dort DHCP
    Port = gZVT2_Port ' Standard eigentlich 22007, aber alle bisher getesteten Geräten haben 22000 eingestellt
    Passwort = "000000" ' Kassiererpasswort
    Protokollpfad = "" ' Wenn nichts angegeben, dann in Eigene Dokumente\GUB\ZVTLOG
    KasseNr = 1 ' für jede Kasse unterschiedlich übergeben, wird im Protokolldateinamen verwendet
    
    'Kassedruck = 0 ' 1 = Kassensoftware druckt Kundenbeleg (nur Professional-Version), 0 = Terminal druckt Kundenbeleg
    If gbZVT2_KBDrucken = True Then
        Kassedruck = 1
    Else
        Kassedruck = 0
    End If
    
    
    Funktion = 0 ' 0 = Zahlen, 1 = Diagnose, 2 = Kassenschnitt, 3 = Storno
    
    Betrag = sBetrag ' Betrag in cent
    
    Test = 0 ' 1 = Testmodus, keine Kommunikation mit dem Terminal
    
    Lizenz = gZVT2_Lizenz ' Lizenzkey passend zur Terminal-ID
    
    Provider = 0 ' 0 = Standardlastschrifttext, 1 = Telecash , 2 = Easycash
    dialog = 2
    
    
    'manuell, dann diese drei Werte
    Kartennummer = sKartennummer
    Kartegueltig = sKartegueltig 'JJMM
    KarteCVC = sKarteCVC

    'Funktion rufen
    Zahlen
    
    
    
    
    If bmitMess = True Then
    
        'Rückgabewerte ausgeben
        MsgBox "Ergebnis: " & Ergebnis & vbCr & _
        ErgebnisText & vbCr & _
        "Ergebnis lang: " & ErgebnisLang & vbCr & _
        "Autorisierungsergebnis: " & Autorisierungsergebnis & vbCr & _
        "Kartentyp: " & Kartentyp & vbCr & _
        "Kartentyp Text: " & KartentypText & vbCr & _
        "Kundenbeleg: " & Kundenbeleg & vbCr & _
        "Haendlerbeleg: " & Haendlerbeleg
    
    End If
    
    
    
    
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT2"
    Fehler.gsFunktion = "Zahlung_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Storno_ZVT2(sBeNR As String, sStornoBetrag As String, bmitMess As Boolean, Optional sKartennummer As String, Optional sKartegueltig As String, Optional sKarteCVC As String) As String
    On Error GoTo LOKAL_ERROR

    Storno_ZVT2 = ""
    
    'Beispielwerte setzen
    COM = "LAN" ' Alternativ "COM" (automatische Com-Port-Erkennung) oder z.B. COM11 (fixer COM-Port)
    ComSpeed = 9600 ' Geräteabhängig, Standard = 9600
    ComStop = 2 ' Geräteabhängig 1 oder 2, Standard = 2
    IP = gZVT2_IP ' wenn IP verwendet wird, dann bitte IP-Adresse am EC-Gerät fest einstellen, Standard ist dort DHCP
    Port = gZVT2_Port ' Standard eigentlich 22007, aber alle bisher getesteten Geräten haben 22000 eingestellt
    
    Passwort = "000000" ' Kassiererpasswort
    Protokollpfad = "" ' Wenn nichts angegeben, dann in Eigene Dokumente\GUB\ZVTLOG
    KasseNr = 1 ' für jede Kasse unterschiedlich übergeben, wird im Protokolldateinamen verwendet
    
'    Kassedruck = 0 ' 1 = Kassensoftware druckt Kundenbeleg (nur Professional-Version), 0 = Terminal druckt Kundenbeleg
    If gbZVT2_KBDrucken = True Then
        Kassedruck = 1
    Else
        Kassedruck = 0
    End If
    
    Funktion = 3 ' 0 = Zahlen, 1 = Diagnose, 2 = Kassenschnitt, 3 = Storno
    StornoBetrag = sStornoBetrag ' Betrag in cent
    StornoBelegNr = sBeNR
    
    Test = 0 ' 1 = Testmodus, keine Kommunikation mit dem Terminal
    Lizenz = gZVT2_Lizenz ' Lizenzkey passend zur Terminal-ID
    Provider = 0 ' 0 = Standardlastschrifttext, 1 = Telecash , 2 = Easycash
    dialog = 2
    
    'manuell, dann diese drei Werte
    Kartennummer = sKartennummer
    Kartegueltig = sKartegueltig 'JJMM
    KarteCVC = sKarteCVC

    'Funktion rufen
    Storno
    
    If bmitMess = True Then

        'Rückgabewerte ausgeben
        MsgBox "Ergebnis: " & Ergebnis & vbCr & _
        ErgebnisText & vbCr & _
        "Ergebnis lang: " & ErgebnisLang & vbCr & _
        "Autorisierungsergebnis: " & Autorisierungsergebnis & vbCr & _
        "Kartentyp: " & Kartentyp & vbCr & _
        "Kartentyp Text: " & KartentypText & vbCr & _
        "Kundenbeleg: " & Kundenbeleg & vbCr & _
        "Haendlerbeleg: " & Haendlerbeleg
    
    
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT2"
    Fehler.gsFunktion = "Storno_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Kassenschnitt_ZVT2(bmitMess As Boolean) As String
    On Error GoTo LOKAL_ERROR

    Kassenschnitt_ZVT2 = ""
    
    'Beispielwerte setzen
    COM = "LAN" ' Alternativ "COM" (automatische Com-Port-Erkennung) oder z.B. COM11 (fixer COM-Port)
    ComSpeed = 9600 ' Geräteabhängig, Standard = 9600
    ComStop = 2 ' Geräteabhängig 1 oder 2, Standard = 2
'    IP = "192.168.1.60" ' wenn IP verwendet wird, dann bitte IP-Adresse am EC-Gerät fest einstellen, Standard ist dort DHCP

    IP = gZVT2_IP

'    Port = 22000 ' Standard eigentlich 22007, aber alle bisher getesteten Geräten haben 22000 eingestellt
    Port = gZVT2_Port
    Passwort = "000000" ' Kassiererpasswort
    Protokollpfad = "" ' Wenn nichts angegeben, dann in Eigene Dokumente\GUB\ZVTLOG
    KasseNr = 1 ' für jede Kasse unterschiedlich übergeben, wird im Protokolldateinamen verwendet
    
    
    'Kassedruck = 0 ' 1 = Kassensoftware druckt Kundenbeleg (nur Professional-Version), 0 = Terminal druckt Kundenbeleg
    If gbZVT2_KBDrucken = True Then
        Kassedruck = 1
    Else
        Kassedruck = 0
    End If
    
    Funktion = 2 ' 0 = Zahlen, 1 = Diagnose, 2 = Kassenschnitt, 3 = Storno
    
'    Betrag = 7 ' Betrag in cent
'    Betrag = sBetrag
    
    Test = 0 ' 1 = Testmodus, keine Kommunikation mit dem Terminal
'    Lizenz = "" ' Lizenzkey passend zur Terminal-ID
    Lizenz = gZVT2_Lizenz
    Provider = 0 ' 0 = Standardlastschrifttext, 1 = Telecash , 2 = Easycash
    dialog = 2

    'Funktion rufen
    Kassenschnitt
    
    If bmitMess = True Then

        'Rückgabewerte ausgeben
        MsgBox "Ergebnis: " & Ergebnis & vbCr & _
        ErgebnisText & vbCr & _
        "Ergebnis lang: " & ErgebnisLang & vbCr & _
        "Autorisierungsergebnis: " & Autorisierungsergebnis & vbCr & _
        "Kartentyp: " & Kartentyp & vbCr & _
        "Kartentyp Text: " & KartentypText & vbCr & _
        "Kundenbeleg: " & Kundenbeleg & vbCr & _
        "Haendlerbeleg: " & Haendlerbeleg
    
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Kassenschnitt_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function BelegWiederholung_ZVT2(bmitMess As Boolean) As String
    On Error GoTo LOKAL_ERROR

    BelegWiederholung_ZVT2 = ""
    
    'Beispielwerte setzen
    COM = "LAN" ' Alternativ "COM" (automatische Com-Port-Erkennung) oder z.B. COM11 (fixer COM-Port)
    ComSpeed = 9600 ' Geräteabhängig, Standard = 9600
    ComStop = 2 ' Geräteabhängig 1 oder 2, Standard = 2
'    IP = "192.168.1.60" ' wenn IP verwendet wird, dann bitte IP-Adresse am EC-Gerät fest einstellen, Standard ist dort DHCP

    IP = gZVT2_IP


'    Port = 22000 ' Standard eigentlich 22007, aber alle bisher getesteten Geräten haben 22000 eingestellt
    Port = gZVT2_Port
    
    
    Passwort = "000000" ' Kassiererpasswort
    Protokollpfad = "" ' Wenn nichts angegeben, dann in Eigene Dokumente\GUB\ZVTLOG
    KasseNr = 1 ' für jede Kasse unterschiedlich übergeben, wird im Protokolldateinamen verwendet
    
'    Kassedruck = 0 ' 1 = Kassensoftware druckt Kundenbeleg (nur Professional-Version), 0 = Terminal druckt Kundenbeleg
    
    If gbZVT2_KBDrucken = True Then
        Kassedruck = 1
    Else
        Kassedruck = 0
    End If
    
    
    Funktion = 5 ' 0 = Zahlen, 1 = Diagnose, 2 = Kassenschnitt, 3 = Storno
    
'    Betrag = 7 ' Betrag in cent
'    Betrag = sBetrag
    
    Test = 0 ' 1 = Testmodus, keine Kommunikation mit dem Terminal
'    Lizenz = "" ' Lizenzkey passend zur Terminal-ID
    Lizenz = gZVT2_Lizenz
    Provider = 0 ' 0 = Standardlastschrifttext, 1 = Telecash , 2 = Easycash
    dialog = 2

    'Funktion rufen
    Belegwiederholung
    
    If bmitMess = True Then

        'Rückgabewerte ausgeben
        MsgBox "Ergebnis: " & Ergebnis & vbCr & _
        ErgebnisText & vbCr & _
        "Ergebnis lang: " & ErgebnisLang & vbCr & _
        "Autorisierungsergebnis: " & Autorisierungsergebnis & vbCr & _
        "Kartentyp: " & Kartentyp & vbCr & _
        "Kartentyp Text: " & KartentypText & vbCr & _
        "Kundenbeleg: " & Kundenbeleg & vbCr & _
        "Haendlerbeleg: " & Haendlerbeleg
    
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "BelegWiederholung_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function letzter_Kundenbeleg_ZVT2() As String
    On Error GoTo LOKAL_ERROR

    letzter_Kundenbeleg_ZVT2 = RegRead("Drucktext")
    


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "letzter_Kundenbeleg_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub lese_ZVT_opt2()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    gZVT2_IP = ""
    gZVT2_Lizenz = ""
    gZVT2_Port = ""
    giZVT2_Fehler = 0
    gbZVT2_KBDrucken = False
    gbZVT2_Kartenwahl = False
    
    gbZVT2_HBDrucken = False
    giZVT2_TIMEOUT = 0
    gsZVT2_VirtuellID = ""
            
    
    If NewTableSuchenDBKombi("ZVTOPT2", gdApp) Then

        Set rsrs = gdApp.OpenRecordset("select * from ZVTOPT2")
        If Not rsrs.EOF Then
            
            If Not IsNull(rsrs!IP) Then
                gZVT2_IP = rsrs!IP
            End If
            
            If Not IsNull(rsrs!Lizenz) Then
                gZVT2_Lizenz = rsrs!Lizenz
            End If
            
            If Not IsNull(rsrs!Port) Then
                gZVT2_Port = rsrs!Port
            End If
            
            If Not IsNull(rsrs!KBDrucken) Then
                gbZVT2_KBDrucken = rsrs!KBDrucken
            End If
            
            If Not IsNull(rsrs!Kartenwahl) Then
                gbZVT2_Kartenwahl = rsrs!Kartenwahl
            End If
            
            If Not IsNull(rsrs!HBDrucken) Then
                gbZVT2_HBDrucken = rsrs!HBDrucken
            End If
            
            If Not IsNull(rsrs!TimeOut) Then
                giZVT2_TIMEOUT = rsrs!TimeOut
            End If
            
            If Not IsNull(rsrs!VIRTUELLEID) Then
                gsZVT2_VirtuellID = rsrs!VIRTUELLEID
            End If
            

            
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "lese_ZVT_opt2"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub





