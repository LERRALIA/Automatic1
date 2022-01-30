VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL210 
   Caption         =   "ZVT 2"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   12465
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox TxtTimeout 
      Height          =   285
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   23
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox TxtVirtuelleID 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   21
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CheckBox CHK_HBDrucken 
      Caption         =   "Händlerbeleg drucken (nur Professionell-Version)"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4200
      Width           =   6975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kreditkarte selbst bestimmen (Kartenauswahl beim Kassieren)"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   6975
   End
   Begin VB.CheckBox Chk_KBDRUCKEN 
      Caption         =   "Kundenbeleg auf dem eigenen Bondrucker ausgeben (nur Professionell-Version)"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   6975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   7080
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox txtLizenz 
      Height          =   285
      Left            =   7080
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1800
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Kassenschnitt"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   7080
      MaxLength       =   15
      TabIndex        =   10
      Top             =   1080
      Width           =   5175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   10200
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Zahlung vornehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtBetrag 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   10200
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Storno"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtBelegNr 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   5520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "letzten Kundenbeleg drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "Timeout in sec:"
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "virtuelle Terminal ID(nur Professionell-Version)"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "Port 22000 (Standard eigentlich 22007)"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "Lizenz"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "IP - Adresse"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Einstellungen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Testbetrieb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Betrag in Cent:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Beleg NR:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmWKL210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 'Zahlung vornehmen
            Zahlung_ZVT2 txtBetrag.Text, True
        Case 1 'Storno vornehmen
            Storno_ZVT2 txtBelegNr.Text, txtBetrag.Text, True
        Case 2 'letzten Kundenbeleg drucken
        
            Dim sBelegtext As String
            Dim iDruckzeilen_count As Integer
            ReDim cDruckZeile(1 To 1) As String
            
            sBelegtext = letzter_Kundenbeleg_ZVT2
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
            
        Case 3
            Unload frmWKL210
        Case 4
            speicher_ZVT2
            lese_ZVT_opt2
        Case 6 'Kassenschnitt
            Kassenschnitt_ZVT2 False
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 Einstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher_ZVT2()
On Error GoTo LOKAL_ERROR

    loeschapp "ZVTOPT2"
    CreateTableT2 "ZVTOPT2", gdApp
    

    sSQL = "Insert into ZVTOPT2 ("
    sSQL = sSQL & "  IP "
    sSQL = sSQL & " ,Lizenz "
    sSQL = sSQL & " ,Port "
    sSQL = sSQL & " ,KBDrucken "
    sSQL = sSQL & " ,Kartenwahl "
    
    sSQL = sSQL & " ,HBDrucken "
    sSQL = sSQL & " ,TimeOut "
    sSQL = sSQL & " ,VIRTUELLEID "
    
    sSQL = sSQL & " ) "
    sSQL = sSQL & " values ( "
    
    
    
    sSQL = sSQL & "'" & txtIP.Text & "'"
    sSQL = sSQL & ",'" & Trim(txtLizenz.Text) & "'"
    sSQL = sSQL & ",'" & Trim(txtPort.Text) & "'"
    
    If Chk_KBDRUCKEN.Value = vbChecked Then
        sSQL = sSQL & " ,true"
    Else
        sSQL = sSQL & " ,false"
    End If
    
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " ,true"
    Else
        sSQL = sSQL & " ,false"
    End If
    
    
    
    If CHK_HBDrucken.Value = vbChecked Then
        sSQL = sSQL & " ,true"
    Else
        sSQL = sSQL & " ,false"
    End If
    
    sSQL = sSQL & "," & Trim(TxtTimeout.Text) & ""
    sSQL = sSQL & ",'" & Trim(TxtVirtuelleID.Text) & "'"
    
    
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError

'    lese_ZVT_opt2

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_ZVT2"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 Einstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    frmWKL116.Top = Screen.Height / 2 - frmWKL116.Height / 2
    frmWKL116.Left = Screen.Width / 2 - frmWKL116.Width / 2
    
    lese_ZVT_opt2

    txtIP.Text = gZVT2_IP
    txtLizenz.Text = gZVT2_Lizenz
    txtPort.Text = gZVT2_Port
    
    If gbZVT2_KBDrucken = True Then
        Chk_KBDRUCKEN.Value = vbChecked
    Else
        Chk_KBDRUCKEN.Value = vbUnchecked
    End If
    
    If gbZVT2_Kartenwahl = True Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    
    
    
    If gbZVT2_HBDrucken = True Then
        CHK_HBDrucken.Value = vbChecked
    Else
        CHK_HBDrucken.Value = vbUnchecked
    End If
    
    TxtTimeout.Text = giZVT2_TIMEOUT
    TxtVirtuelleID.Text = gsZVT2_VirtuellID
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil ZVT2 Einstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Function Beispielzahlung() As String

    'Beispielwerte setzen
    COM = "LAN" ' Alternativ "COM" (automatische Com-Port-Erkennung) oder z.B. COM11 (fixer COM-Port)
    ComSpeed = 9600 ' Geräteabhängig, Standard = 9600
    ComStop = 2 ' Geräteabhängig 1 oder 2, Standard = 2
    IP = "192.168.1.60" ' wenn IP verwendet wird, dann bitte IP-Adresse am EC-Gerät fest einstellen, Standard ist dort DHCP
    Port = 22000 ' Standard eigentlich 22007, aber alle bisher getesteten Geräten haben 22000 eingestellt
    Passwort = "000000" ' Kassiererpasswort
    Protokollpfad = "" ' Wenn nichts angegeben, dann in Eigene Dokumente\GUB\ZVTLOG
    KasseNr = 1 ' für jede Kasse unterschiedlich übergeben, wird im Protokolldateinamen verwendet
    Kassedruck = 1 ' 1 = Kassensoftware druckt Kundenbeleg (nur Professional-Version), 0 = Terminal druckt Kundenbeleg
    Funktion = 0 ' 0 = Zahlen, 1 = Diagnose, 2 = Kassenschnitt, 3 = Storno
    Betrag = 7 ' Betrag in cent
    Test = 0 ' 1 = Testmodus, keine Kommunikation mit dem Terminal
    Lizenz = "" ' Lizenzkey passend zur Terminal-ID
    Provider = 0 ' 0 = Standardlastschrifttext, 1 = Telecash , 2 = Easycash
    dialog = 1

    'Funktion rufen
    Zahlen

    'Rückgabewerte ausgeben
    MsgBox "Ergebnis: " & Ergebnis & vbCr & _
    ErgebnisText & vbCr & _
    "Ergebnis lang: " & ErgebnisLang & vbCr & _
    "Autorisierungsergebnis: " & Autorisierungsergebnis & vbCr & _
    "Kartentyp: " & Kartentyp & vbCr & _
    "Kartentyp Text: " & KartentypText & vbCr & _
    "Kundenbeleg: " & Kundenbeleg & vbCr & _
    "Haendlerbeleg: " & Haendlerbeleg

    Beispielzahlung = "Fertig"
    
End Function



