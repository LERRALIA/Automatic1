VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWK20b 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "offene Kassenvorg‰nge"
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   1485
   ClientWidth     =   10905
   Icon            =   "frmWK20b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6810
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   120
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   8
      Top             =   3960
      Width           =   5295
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C000&
      Caption         =   "alle Anzeigen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   4
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
   Begin MSComctlLib.TreeView List3 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4683
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9120
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9120
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Caption         =   "Ausw‰hlen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10695
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9120
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "‹bersicht drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdEinzel 
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Einzelne ausw‰hlen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9120
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Angebot drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label15 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   7455
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "offene Kassenvorg‰nge"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   10800
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmWK20b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check9_Click()
On Error GoTo LOKAL_ERROR
    
    zeigvorgaenge
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check9_Click"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub cmdEinzel_Click()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cLBSatz As String
    Dim cLfnr As String
    Dim sZeilennummer As String
    Dim bFound As Boolean
    Dim sZeilNr() As String
    Dim lAnzZeile As Long
    
    bFound = False
    lAnzZeile = 0
    ReDim sZeilNr(1 To 1) As String

    If List2.ListCount = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount

    If Not bFound Then
        anzeigeNew "rot", "Bitte einen Artikel in der Liste ausw‰hlen!", Label15
        Exit Sub
    End If
    
    If bFound = True Then
        
    
        cLfnr = Trim(Text1.Text)
    
        For lcount = 0 To List2.ListCount - 1
            If List2.Selected(lcount) = True Then
                sZeilennummer = Trim(Right(List2.list(lcount), 3))
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve sZeilNr(1 To lAnzZeile) As String
                sZeilNr(lAnzZeile) = sZeilennummer
            
                
                
            End If
        Next lcount
        
        If IsNumeric(cLfnr) Then
            HoleUnterbrochenenBonWK20b_einzelneArtikel cLfnr, sZeilNr
            
            Unload frmWK20b
        End If
    

    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEinzel_Click"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    voreinstellungspeichernE20D
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cBedNr As String
    Dim cBedNrBon As String
    
    Select Case Index
        Case Is = 0

            cLBSatz = Trim(Text1.Text)
            
            If IsNumeric(cLBSatz) Then
                HoleUnterbrochenenBonWK20b cLBSatz
                Unload frmWK20b
            Else
                anzeigeNew "rot", "Bitte einen Eintrag in der Liste ausw‰hlen!", Label15
            End If
                
        Case Is = 1
            Unload frmWK20b
            
        Case Is = 2
            drucke_offene_Kassenvorg‰nge
            
        Case Is = 3
            cLBSatz = Trim(Text1.Text)
            
            If IsNumeric(cLBSatz) Then
                drucke_Angebot cLBSatz, "Angebot", True 'gbDINA4RECHFU
            End If
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub drucke_Angebot(cLfdNr As String, cZahlart As String, bReFuss As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lBedNr              As Long
    Dim cBedname            As String
    Dim cKundnr             As String
    Dim lAktsatz            As Long
    Dim rs                  As DAO.Recordset
    Dim rsrs                As DAO.Recordset
    Dim dUSTV               As Double
    Dim dUSTE               As Double
    Dim cFirma              As String
    Dim cVname              As String
    Dim cNName              As String
    Dim cTitel              As String
    Dim cPlz                As String
    Dim cStadt              As String
    Dim cStrasse            As String
    Dim cAnrede             As String
    Dim cPrintFiTiAnVoNa    As String
    Dim sSQL                As String
    Dim cLBSatz             As String
    
    loeschNEW "ANGEBOTNOW", gdBase
    CreateTable "ANGEBOTNOW", gdBase
    loeschNEW "DAGKOPF", gdBase
    CreateTable "DAGKOPF", gdBase

    sSQL = "Select * from BONPAUSE where LFDNR = " & cLfdNr & " order by LBZEILE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNR) Then
            lBedNr = rsrs!BEDNR
        Else
            lBedNr = 0
        End If

        If Not IsNull(rsrs!KdNr) Then
            cKundnr = rsrs!KdNr
        Else
            cKundnr = ""
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cBedname = ermBEDbez(lBedNr)
    
    Dim sMenge As String
    Dim sArtnr As String
    Dim sBezeich As String
    Dim sMWST As String
    Dim sVKPR As String
    Dim sKVKPR As String
    Dim sMopreis As String
    Dim sGPreis As String
    
    
    sSQL = "Select * from ANGEBOTNOW "
    FnOpenrecordset rs, sSQL, 1, gdBase
    
    sSQL = "Select * from BONPAUSE where LFDNR = " & cLfdNr & " order by LBZEILE "
    FnOpenrecordset rsrs, sSQL, 1, gdBase
    lAktsatz = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAktsatz = lAktsatz + 1
            
            If Not IsNull(rsrs!lbtext) Then
                cLBSatz = rsrs!lbtext
            End If
            
'            MsgBox "Menge = " & Mid(cLBSatz, 1, 5)
'            MsgBox "ArtNr = " & Mid(cLBSatz, 7, 6)
'            MsgBox "Bezeich = " & Mid(cLBSatz, 14, 35)
'            MsgBox "EPreis nach Rabatt = " & Mid(cLBSatz, 50, 9)
'            MsgBox "GPreis = " & Mid(cLBSatz, 60, 9)
'            MsgBox "MWST-Kz = " & Mid(cLBSatz, 72, 1)
'            MsgBox "Listenpreis/Sonderpreis = " & Mid(cLBSatz, 74, 9)
'            MsgBox "Betrag ArtRabatt = " & Mid(cLBSatz, 84, 9)
'            MsgBox "erzielter VK-Preis = " & Mid(cLBSatz, 94, 9)
'            MsgBox "Betrag volle MWST = " & Mid(cLBSatz, 104, 9)
'            MsgBox "Betrag erm. MWST = " & Mid(cLBSatz, 114, 9)
'            MsgBox "ArtRabatt % = " & Mid(cLBSatz, 124, 3)
'            MsgBox "Listen-VK = " & Mid(cLBSatz, 128, 9)
'            MsgBox "Restmenge = " & Mid(cLBSatz, 138, 9)
            
            sMenge = Trim(Mid(cLBSatz, 1, 5))
            sArtnr = Trim(Mid(cLBSatz, 7, 6))
            sBezeich = Trim(Mid(cLBSatz, 14, 35))
            sMWST = Trim(Mid(cLBSatz, 72, 1))
            sGPreis = Trim(Mid(cLBSatz, 60, 9))
            
            sMopreis = Trim(Mid(cLBSatz, 177, 8))
            If sMopreis = "" Then sMopreis = "0"
            
            sVKPR = Trim(Mid(cLBSatz, 128, 9))
            
            rs.AddNew
            rs!artnr = sArtnr
            rs!BEZEICH = sBezeich
            
            If Val(sMenge) <> 0 Then
                rs!vkpr = sGPreis / sMenge
            Else
                rs!vkpr = sGPreis
            End If
            
            rs!ANZAHL = sMenge
            rs!MWST = sMWST
            rs!KVKPR1 = sVKPR
            rs!lfnr = 0
            rs!posinr = lAktsatz
            rs!BEDNR = lBedNr
            rs!MOPREIS = sMopreis

            rs.Update
            
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    rs.Close: Set rs = Nothing
    
    
    
    
    
    
    
    
    
    cKundnr = Trim(cKundnr)
    
    If cKundnr = "" Then
        cKundnr = "0"
    End If

    dUSTV = ermUstv("ANGEBOTNOW")
    dUSTE = ermUste("ANGEBOTNOW")

    
    If cKundnr <> "0" Then
        
        cFirma = lookingForKundendaten(Trim(cKundnr)).firma
        cVname = lookingForKundendaten(Trim(cKundnr)).vorname
        cNName = lookingForKundendaten(Trim(cKundnr)).nachname
        cTitel = lookingForKundendaten(Trim(cKundnr)).titel
        cPlz = lookingForKundendaten(Trim(cKundnr)).Plz
        cStadt = lookingForKundendaten(Trim(cKundnr)).Ort
        cStrasse = lookingForKundendaten(Trim(cKundnr)).strasse
        cAnrede = lookingForKundendaten(Trim(cKundnr)).anrede
        
        cPrintFiTiAnVoNa = ""
        
        If cFirma <> "" Then
            cPrintFiTiAnVoNa = cFirma
        End If
        
        cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & vbCrLf
        
        If cAnrede <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cAnrede & Space(1)
        End If
        
        If cTitel <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cTitel & Space(1)
        End If
        
        If cVname <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cVname & Space(1)
        End If
        
        If cNName <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cNName
        End If
        
        sSQL = "Insert into DAGKOPF (kundnr,PrintFiTiAnVoNa,name,vorname,titel,plz,stadt,strasse,Firma,anrede,datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        sSQL = sSQL & cKundnr
        sSQL = sSQL & ", '" & cPrintFiTiAnVoNa & "' "
        sSQL = sSQL & ", '" & cNName & "' "
        sSQL = sSQL & ", '" & cVname & "' "
        sSQL = sSQL & ", '" & cTitel & "' "
        sSQL = sSQL & ", '" & cPlz & "' "
        sSQL = sSQL & ", '" & cStadt & "' "
        sSQL = sSQL & ", '" & cStrasse & "' "
        sSQL = sSQL & ", '" & cFirma & "' "
        sSQL = sSQL & ", '" & cAnrede & "' "
        sSQL = sSQL & " , '" & cZahlart & "' "
        sSQL = sSQL & " , '" & cBedname & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
        
    Else
    
        sSQL = "Insert into DAGKOPF (datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        sSQL = sSQL & "  '' "
        sSQL = sSQL & " , '" & cBedname & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    
    
    
    
    
    
    
    Dim cFirmName       As String
    Dim cFirmAdress     As String
    Dim cFirmBank       As String
    Dim cFirmKomm       As String
    Dim cSteuernr       As String
    Dim cKommentar      As String
    
    
    loeschNEW "REFUSS", gdBase
    CreateTableT2 "REFUSS", gdBase
    
    If bReFuss Then
        
        sSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
            cSteuernr = ""
        End If
        
        sSQL = "Insert into REFUSS ( "
        sSQL = sSQL & " STEUERNR"
        sSQL = sSQL & ", FIRMNAME"
        sSQL = sSQL & ", FIRMADRESS"
        sSQL = sSQL & ", FIRMBANK"
        sSQL = sSQL & ", FIRMKOMM"
        sSQL = sSQL & ") values ("
        sSQL = sSQL & " '" & cSteuernr & "'"
        sSQL = sSQL & ", '" & cFirmName & "'"
        sSQL = sSQL & ", '" & cFirmAdress & "'"
        sSQL = sSQL & ", '" & cFirmBank & "'"
        sSQL = sSQL & ", '" & cFirmKomm & "'"
        sSQL = sSQL & ") "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    reportbildschirm "alr", "angebot_mf"

    setzedrucker gcBonDrucker

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucke_Angebot"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub drucke_offene_Kassenvorg‰nge()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "PRINT_BONPAUSE", gdBase
    CreateTableT2 "PRINT_BONPAUSE", gdBase
    
    cSQL = "Insert into PRINT_BONPAUSE Select distinct LFDNR, BEDNR, KDNR, KDNAME, ZSUM , Adate, azeit from BONPAUSE "
    If Check9.Value = vbUnchecked Then
        cSQL = cSQL & " where bednr = '" & Trim(frmWKL20.Text1(0).Text) & "'"
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    reportbildschirm "", "aWKL20b"
    
    setzedrucker gcBonDrucker

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucke_offene_Kassenvorg‰nge"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub zeigvorgaenge()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim iFeld As Integer
    Dim cLBSatz As String
    Dim lcount As Long
    
    List1.Clear

    List3.Nodes.Clear
    List1.AddItem "Bed     KundNr KundenName                       Zwischensumme  Datum    Uhrzeit  V-Nr"
    
    cSQL = "Select distinct LFDNR, BEDNR, KDNR, KDNAME, ZSUM , Adate, azeit from BONPAUSE "
    If Check9.Value = vbUnchecked Then
        cSQL = cSQL & " where bednr = '" & Trim(frmWKL20.Text1(0).Text) & "'"
    End If
  
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!BEDNR) Then
                cFeld = rsrs!BEDNR
            Else
                cFeld = ""
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!KdNr) Then
                cFeld = rsrs!KdNr
            Else
                cFeld = ""
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!KdName) Then
                cFeld = rsrs!KdName
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!ZSUM) Then
                cFeld = rsrs!ZSUM
            Else
                cFeld = ""
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            
            
            If Not IsNull(rsrs!Adate) Then
                cFeld = Format(rsrs!Adate, "DD.MM.YY")
            Else
                cFeld = ""
            End If

            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!AZEIT) Then
                cFeld = Format(rsrs!AZEIT, "HH:MM:SS")
            Else
                cFeld = ""
            End If

            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!LFDNR) Then
                iFeld = rsrs!LFDNR
            Else
                iFeld = 0
            End If
            cFeld = Trim$(Str$(iFeld))
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            

            List3.Nodes.Add Text:=cLBSatz
            
            List3.Nodes(List3.Nodes.Count).BackColor = vbYellow
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lcount = 1 Then
        List3.Nodes(1).Selected = True
        List3_NodeClick List3.Nodes(1)
        List3.Nodes(1).BackColor = vbBlue
        List3.Nodes(1).ForeColor = vbWhite
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigvorgaenge"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub zeigeParkvorg‰nge(cNum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim iFeld As Integer
    Dim cLBSatz As String
    
    List2.Clear

    cSQL = "Select lbtext,LBZEILE  from BONPAUSE where lfdnr = " & cNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!lbtext) Then
                cFeld = rsrs!lbtext
            Else
                cFeld = ""
            End If
            
            cLBSatz = cFeld
            
            If Not IsNull(rsrs!LBZEILE) Then
                cFeld = rsrs!LBZEILE
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & Space(10) & cFeld
            
            List2.AddItem cLBSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigeParkvorg‰nge"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    If NewTableSuchenDBKombi("E20D", gdApp) Then
        voreinstellungladenE20D
    End If
    
    zeigvorgaenge
    anzeigeNew "normal", "V-Nr eingeben oder einen Eintrag in der Liste ausw‰hlen!", Label15

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE20D()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim bo0 As Integer
    
    Set rs = gdApp.OpenRecordset("E20D")
    If Not rs.EOF Then
        If rs!bo0 = True Then
            Check9.Value = vbUnchecked
        Else
            Check9.Value = vbChecked
        End If
    End If
    rs.Close: Set rs = Nothing
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE20D"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE20D()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo0 As Integer
    
    
    loeschNEW "E20D", gdApp
    CreateTable "E20D", gdApp
    
    If Check9.Value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    sSQL = "Insert into E20D ( bo0) "
    sSQL = sSQL & " values (" & bo0 & ")"

    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE20D"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_GotFocus()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    If List3.SelectedItem Is Nothing Then
        Exit Sub
    Else
        For i = 1 To List3.Nodes.Count
            If List3.Nodes(i).Selected = True Then

                Text1.Text = Right(List3.Nodes(i), 4)
                Text1.Refresh
                If Text1.Text <> "" Then
                    zeigeParkvorg‰nge Text1.Text
                End If
                Exit For
            Else

            End If

        Next i
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub List3_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    If List3.SelectedItem Is Nothing Then
        Exit Sub
    Else
        For i = 1 To List3.Nodes.Count
            If List3.Nodes(i).Selected = True Then
            
                Text1.Text = Right(List3.Nodes(i), 4)
                Text1.Refresh
                
                If Text1.Text <> "" Then
                    zeigeParkvorg‰nge Text1.Text
                End If
                Exit For
            Else

            End If
        Next i
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_NodeClick"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
   
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    
    cValid = gcNUM
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Beep
    End If
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command1_Click 0
    ElseIf KeyCode = vbKeyEscape Then
        Command1_Click 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorg‰nge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
