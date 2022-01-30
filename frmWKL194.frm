VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL194 
   BackColor       =   &H00C0C000&
   Caption         =   "neue Coupon - Daten"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL194.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11160
      Top             =   240
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   10320
      Pattern         =   "MASTER!.*"
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   7680
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   0
      Top             =   7080
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
      Caption         =   "Einlesen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Coupon-Regeln:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Coupon-Daten:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Möchten Sie diese übernehmen, so klicken Sie auf ""Einlesen""."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neue Coupon-Regeln stehen bereit. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "neue Coupon-Daten"
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
      Width           =   11535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL194"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glartv      As Long
Dim glartb      As Long
Dim iSec        As Integer

Private Function wert_back(cSatz As String, cTag As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lItemPos        As Long
    Dim lItemEndePos    As Long
    
    wert_back = ""
    
    lItemPos = InStr(1, cSatz, "<" & cTag & ">")
    lItemEndePos = InStr(lItemPos, cSatz, "</" & cTag & ">")
    
    lItemPos = lItemPos + Len(cTag) + 2
    wert_back = Mid(cSatz, lItemPos, lItemEndePos - lItemPos)

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "wert_back"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub Speicher_EAN(cSatz As String, cCID As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cEAN As String
    
    Dim lItemPos        As Long
    Dim lItemEndePos    As Long
    
    cSatz = SwapStr(cSatz, "<ean>", "")
    cSatz = SwapStr(cSatz, Chr(10), "")
    cSatz = SwapStr(cSatz, Chr(13), "")
    cSatz = SwapStr(cSatz, " ", "")
    cSatz = SwapStr(cSatz, "<", "")
    cSatz = SwapStr(cSatz, "ean>", "")
    
    cSatz = Left(cSatz, Len(cSatz) - 1)
    
    
    Dim sArray() As String
    sArray = Split(cSatz, "/")

    For i = 0 To UBound(sArray)

        cEAN = sArray(i)
        
        sSQL = "Insert into COUPONPRODUKTE (COUPON_ID,EAN) values ("
        sSQL = sSQL & cCID
        sSQL = sSQL & ", '" & cEAN & "'"
        sSQL = sSQL & ")"
        gdBase.Execute sSQL, dbFailOnError
        
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Speicher_EAN"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Speicher_EAN_PlusArtikel(cSatz As String, cCID As String, cBez As String, cCent As String, sHerst As String _
, sStart As String, sEnd As String, sCoupontype As String, sMultiredemption As String, sQuantity As String, sProductamount As String _
, sTender As String)

    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cEAN As String
    Dim cNeueEAN As String
    Dim cZeichen As String
    Dim rsR As DAO.Recordset
    
    Dim lItemPos        As Long
    Dim lItemEndePos    As Long
    
    cSatz = SwapStr(cSatz, "<ean>", "")
    cSatz = SwapStr(cSatz, Chr(10), "")
    cSatz = SwapStr(cSatz, Chr(13), "")
    cSatz = SwapStr(cSatz, " ", "")
    cSatz = SwapStr(cSatz, "<", "")
    cSatz = SwapStr(cSatz, "ean>", "")
    
    cSatz = Left(cSatz, Len(cSatz) - 1)
    
    
    Dim sArray() As String
    sArray = Split(cSatz, "/")

    For i = 0 To UBound(sArray)

        cEAN = Trim(sArray(i))
        
        cNeueEAN = ""
        For j = 1 To Len(cEAN)
            cZeichen = Mid(cEAN, j, 1)
            If IsNumeric(cZeichen) = True Then
                cNeueEAN = cNeueEAN & cZeichen
            End If
        Next j
        
        Coupon_Artikel_anlegen cNeueEAN, cCID, Left(cBez, 35), cCent
        
        Set rsR = gdBase.OpenRecordset("COUPONREGELN")
        rsR.AddNew
        rsR!COUPON_ID = cCID
        rsR!COUPON_HERSTELLER = sHerst
        rsR!COUPON_TITEL = cBez
        rsR!COUPON_STARTDATUM = sStart
        rsR!COUPON_ENDEDATUM = sEnd
        rsR!COUPON_TYPE = sCoupontype
        rsR!COUPON_MEHRFACHEINLÖSUNG = sMultiredemption
        rsR!COUPON_MINDESTANZAHL = sQuantity
        rsR!COUPON_MINDESTUMSATZ = sProductamount
        rsR!COUPON_DISCOUNTPREIS = cCent
        rsR!COUPON_TENDER = sTender
        rsR!COUPON_ARTNR = 0
        rsR!COUPON_EAN = cNeueEAN
        rsR.Update
        rsR.Close: Set rsR = Nothing
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Speicher_EAN_PlusArtikel"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Coupon_Artikel_anlegen(cEAN As String, cCID As String, cBez As String, cCent As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim dPreis As Double
    dPreis = (CDbl(cCent) / 100) * -1
    
    cBez = SwapStr(cBez, "'", "")
    
    sSQL = "Update COUPONREGELN set COUPON_EAN = '' where COUPON_EAN  = '" & cEAN & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into COUPON_NEUE "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & ", bezeich "
    sSQL = sSQL & ", KVKPR1  "
    sSQL = sSQL & ", EAN  "
    sSQL = sSQL & ", NOTIZEN "
    sSQL = sSQL & " ) values ( "
    sSQL = sSQL & 0
    sSQL = sSQL & ", '" & cBez & "'"
    sSQL = sSQL & ", '" & dPreis & "'"
    sSQL = sSQL & ", '" & cEAN & "'"
    sSQL = sSQL & ", '" & cCID & "'"
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Coupon_Artikel_anlegen"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub CouponRegelnEinlesen(sPfad As String, sDatei As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL                As String
    Dim rsR                 As Recordset
    Dim iFileNr             As Integer
    Dim cSatz1              As String
    Dim lPos                As Long
    Dim cEinzelsatz         As String
    Dim cProduktEANsatz     As String
    Dim cCouponEANsatz      As String
    Dim lLenfil             As Long
    Dim lItemPos            As Long
    Dim lItemEndePos        As Long
    Dim sCouponID           As String
    Dim sArtikelbez         As String
    Dim sDiscountPrice      As String
    Dim lreservArtnr        As Long
    
    
    
    Dim sHerst As String
    Dim sStart As String
    Dim sEnd As String
    Dim sCoupontype As String
    Dim sMultiredemption As String
    Dim sQuantity As String
    Dim sProductamount As String
    Dim sTender As String
    
    
    
'    Label1(6).Visible = True
'    Label1(10).Visible = True

    Screen.MousePointer = 11
    
    loeschNEW "COUPON_NEUE", gdBase
    CreateTableT2 "COUPON_NEUE", gdBase
    
    If NewTableSuchenDBKombi("COUPONREGELN", gdBase) = False Then
        CreateTableT2 "COUPONREGELN", gdBase
    End If
    
    If NewTableSuchenDBKombi("COUPONPRODUKTE", gdBase) = False Then
        CreateTableT2 "COUPONPRODUKTE", gdBase
    End If
    
    CheckIndex "COUPONPRODUKTE", "COUPON_ID", "", gdBase
    CheckIndex "COUPONREGELN", "COUPON_ID", "", gdBase
    
    lPos = 1
    
    iFileNr = FreeFile

    Open sPfad & "\" & sDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1
        lLenfil = Len(cSatz1)
        
        Do
            lItemPos = InStr(lPos, cSatz1, "<couponrule>")
            If lItemPos = 0 Then Exit Do
            lItemEndePos = InStr(lItemPos, cSatz1, "</couponrule>")
            
            lPos = lItemEndePos
            
            cEinzelsatz = Mid(cSatz1, lItemPos, lItemEndePos - lItemPos)

            sCouponID = wert_back(cEinzelsatz, "id")
            
            cSQL = "Delete * from COUPONPRODUKTE where COUPON_ID = " & sCouponID
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Delete * from COUPONREGELN where COUPON_ID = " & sCouponID
            gdBase.Execute cSQL, dbFailOnError
            
            sHerst = wert_back(cEinzelsatz, "manufacturer")
            sHerst = SwapStr(sHerst, "&amp;", "&")
            sArtikelbez = wert_back(cEinzelsatz, "title")
            sArtikelbez = SwapStr(sArtikelbez, "&amp;", "&")
            sStart = wert_back(cEinzelsatz, "startdate")
            sEnd = wert_back(cEinzelsatz, "enddate")
            sCoupontype = wert_back(cEinzelsatz, "coupontype")
            sMultiredemption = wert_back(cEinzelsatz, "multiredemption")
            sQuantity = wert_back(cEinzelsatz, "quantity")
            sProductamount = wert_back(cEinzelsatz, "productamount")
            sDiscountPrice = wert_back(cEinzelsatz, "discountprice")
            sTender = wert_back(cEinzelsatz, "tender")
            
            cProduktEANsatz = wert_back(cEinzelsatz, "products")
            cCouponEANsatz = wert_back(cEinzelsatz, "coupons")
            
            
            Speicher_EAN cProduktEANsatz, sCouponID
            
            
            Speicher_EAN_PlusArtikel cCouponEANsatz, sCouponID, sArtikelbez, sDiscountPrice, sHerst, sStart, sEnd, sCoupontype _
            , sMultiredemption, sQuantity, sProductamount, sTender
          
            anzeige "normal", lPos & " von " & lLenfil, Label1(4)
            
            
        Loop While lLenfil >= lPos
    End If
    
    Close iFileNr
    
    anzeige "normal", "Für neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)

    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)

    sSQL = "Select * from COUPON_NEUE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!artnr = lreservArtnr
            rsrs.Update

            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verfügung (Einstellungen überprüfen).", Label1(4)
                Exit Sub
            Else
                lreservArtnr = lvergebeArtnr
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    If lvergebeArtnr > 0 Then
        sSQL = "Update FFE set ARTNRV = " & lvergebeArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If Not NewTableSuchenDBKombi("COUPE", gdBase) Then 'das erste Mal
        CreateTableT2 "COUPE", gdBase
        sSQL = "Insert into COUPE (linr) values (0)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    Dim lLinr As Long
    
    lLinr = checkCouponinLISRT()

    Uebernahme_Coupon lLinr
    
    
    
    sSQL = "Update COUPONREGELN inner join COUPON_NEUE "
    sSQL = sSQL & " on COUPONREGELN.COUPON_EAN = COUPON_NEUE.EAN "
    sSQL = sSQL & " set COUPONREGELN.COUPON_ARTNR = COUPON_NEUE.ARTNR "
    sSQL = sSQL & " where COUPONREGELN.COUPON_ARTNR = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    


    Kill sPfad & "\" & sDatei

    anzeige "normal", "Fertig! Couponübernahme beendet", Label1(4)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CouponRegelnEinlesen"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function checkCouponinLISRT() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsLi As Recordset
    
    checkCouponinLISRT = 0

    sSQL = "Select linr from COUPE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!linr) Then
            checkCouponinLISRT = rsrs!linr
            
            sSQL = "Select * from LISRT where LINR = " & checkCouponinLISRT
            sSQL = sSQL & " and ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
            Set rsLi = gdBase.OpenRecordset(sSQL)
            If rsLi.RecordCount = 0 Then
                checkCouponinLISRT = 0
            End If
            rsLi.Close
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If checkCouponinLISRT = 0 Then
    
        Screen.MousePointer = 0
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
        If gF2Prompt.cFeld <> "" Then
            gsAnzeige00a = "Bitte den Coupon - Lieferant auswählen!"
            frmWK00a.Show 1
        End If
        gsAnzeige00a = ""
        
        anzeige "normal", "Der Lieferant: " & gF2Prompt.cWahl & " wurde zugeordnet.", Label1(4)
        
        If gF2Prompt.cWahl <> "" Then
             checkCouponinLISRT = CDbl(gF2Prompt.cWahl)
        End If
        
        If checkCouponinLISRT <> 0 Then
            sSQL = "update COUPE set linr = " & checkCouponinLISRT
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkCouponinLISRT"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub CouponIDlöschen(sNotizen As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim sArtnr As String
    
    cSQL = "Select * from Artikel where Notizen = '" & sNotizen & "'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = Trim(rsrs!artnr)
            End If
            
            cSQL = "Delete from Artikel where artnr = " & sArtnr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Delete from Artlief where artnr = " & sArtnr
            gdBase.Execute cSQL, dbFailOnError

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CouponIDlöschen"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SicherheitslöschenCoupon(sArtnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from Artikel where artnr = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Artlief where artnr = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SicherheitslöschenCoupon"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub COUPON_EAN_ENTFERNEN(sEAN As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Update Artikel set EAN = '' where EAN = '" & sEAN & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN2 = '' where EAN2 = '" & sEAN & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN3 = '' where EAN3 = '" & sEAN & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artean_k set EAN = '' where EAN = '" & sEAN & "'"
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "COUPON_EAN_ENTFERNEN"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Uebernahme_Coupon(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim sArtnr          As String
    Dim sNotizen        As String
    Dim sEAN            As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz übernommen...", Label1(4)
    
    'neue Artikel
    cSQL = "Select * from COUPON_NEUE "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = Trim(rsrs!artnr)
            End If
            
            If Not IsNull(rsrs!NOTIZEN) Then
                sNotizen = Trim(rsrs!NOTIZEN)
            End If
            
            If Not IsNull(rsrs!EAN) Then
                sEAN = Trim(rsrs!EAN)
            End If
            
            SicherheitslöschenCoupon sArtnr
            
            CouponIDlöschen sNotizen
            
            COUPON_EAN_ENTFERNEN sEAN

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz übernommen(3)...", Label1(4)
    
    cSQL = "Insert into Artikel Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", 0 as AGN "
    cSQL = cSQL & ", 0 as PGN "
    cSQL = cSQL & ", 0 as LEKPR "
    cSQL = cSQL & ", 0 as VKPR "
    cSQL = cSQL & ", 'V' as MWST "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & ", 'N' as RABATT_OK "
    cSQL = cSQL & ", 'J' as GEFUEHRT "
    cSQL = cSQL & ", 0 as EKPR "
    cSQL = cSQL & ", 'J' as PREISSCHU "
    cSQL = cSQL & ", 'N' as BONUS_OK "
    cSQL = cSQL & ", 'J' as UMS_OK "
    cSQL = cSQL & ", '0' as AWM "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as AUFDAT "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", NOTIZEN "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & " from COUPON_NEUE "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz übernommen(4)...", Label1(4)
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", " & lLinr & " as LINR "
    cSQL = cSQL & ", 0 as LEKPR "
    cSQL = cSQL & ", '' as LIBESNR "
    cSQL = cSQL & ", 0 as MINMEN "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & " from COUPON_NEUE "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz übernommen(5)...", Label1(4)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme_Coupon"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Sicherheitslöschen_mitLinr(sArtnr As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from artlief where artnr = " & sArtnr & " and Linr = " & sLinr
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sicherheitslöschen_mitLinr"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub




Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
 
    Dim sdatname    As String
    Dim i           As Integer
    Dim sSQL        As String
    Dim lLFNR       As Long
    Dim cLfnr       As String
    Dim rsrs        As DAO.Recordset
 
    Select Case Index
        Case 0
            Unload frmWKL194
        Case 1      'Rewe Stammdaten einlesen
        
            Timer1.Enabled = False
            Command5(1).Enabled = False
        
            'Ablaufprotokoll füllen
            'Etiketten erstellen
            'dem Anwender ein Übernahmeergebnis zeigen
            
            CreateTableT2 "CORDER", gdBase
            
            File1.Path = gsKinPfad 'Standard In Pfad
            File1.Pattern = "*.xml"
            File1.Refresh
            
            If File1.ListCount > 0 Then
                'Datei/en stehen an
                For i = 0 To File1.ListCount - 1
                    sdatname = File1.list(i)
                    cLfnr = Mid(sdatname, 11, 3)
                    lLFNR = Val(cLfnr)
                    
                    sSQL = "Insert into CORDER (lfnr,DATNAME)"
                    sSQL = sSQL & " Values ( "
                    sSQL = sSQL & " " & lLFNR & " "
                    sSQL = sSQL & ", '" & sdatname & "' "
                    sSQL = sSQL & " ) "
                    gdBase.Execute sSQL, dbFailOnError
                Next i
            End If
            
            sSQL = "Select * from CORDER order by lfnr asc"
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                    If Not IsNull(rsrs!Datname) Then
                        sdatname = rsrs!Datname
                        
                        CouponRegelnEinlesen gsKinPfad, sdatname
                        
                    End If
                
                rsrs.MoveNext
                Loop
            End If
            rsrs.Close: Set rsrs = Nothing
            
            loeschNEW "CORDER", gdBase
    End Select
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    lesenEinstellungen
    iSec = 0
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lesenEinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    glartv = 600000
    glartb = 700000
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!ARTNRV) Then
                glartv = rsrs!ARTNRV
            Else
                glartv = 600000
            End If
            
            If Not IsNull(rsrs!ARTNRB) Then
                glartb = rsrs!ARTNRB
            Else
                glartb = 700000
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesenEinstellungen"
    Fehler.gsFehlertext = "Im Programmteil Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
'    loeschNEW "STADAPROBELA", gdBase
    
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


Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    iSec = iSec + 1
    
    If iSec >= 10 Then
        Unload frmWKL194
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil neue Coupondaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

