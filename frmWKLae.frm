VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKLae 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   6735
   ClientLeft      =   1770
   ClientTop       =   1485
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6735
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6840
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
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
      Height          =   480
      Left            =   7800
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   6240
      Width           =   1695
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
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
      Caption         =   "Info"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   11
      Cols            =   13
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "festgesetzter Mindestbestand in dieser Filiale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stück(e) in Bestellung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   39
      Left            =   1560
      TabIndex        =   8
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   7
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "letzte Bestellung:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   45
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   8520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "JOOP CALIENTE FUENTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Warenbestand in Filialen "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmWKLae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 0
        Unload frmWKLae
    Case 1
        DruckenFB
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    gcArtNrFiliale = Label5(0).Caption

    frmWKLam.Show 1
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckenFB()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim cDatum      As String
    Dim czeit       As String
    Dim cArtNr      As String
    Dim cEAN        As String
    Dim cBezeich    As String
    Dim i           As Integer
    ReDim cZeilen(0 To giAnzFil + 8) As String
    
    cArtNr = Label5(0).Caption
    cBezeich = Label5(1).Caption
    cEAN = Label5(2).Caption
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    'Drucke den Beleg

    cZeilen(0) = "Filialbestände, gedruckt in: " & gcFilNr
    cZeilen(1) = "-------------------------------"
    cZeilen(2) = "Artikel/EAN"
    cZeilen(3) = cArtNr & "/" & cEAN
    cZeilen(4) = cBezeich
    cZeilen(5) = "Datum: " & cDatum
    cZeilen(6) = "Zeit:  " & czeit
    cZeilen(7) = vbCrLf
    
    
    
    For i = 0 To giAnzFil
        cZeilen(i + 8) = Left(List2.list(i), 27)
    Next i
    
    
    
    DruckeArbeitszeitBelegWK20d cZeilen(), giAnzFil + 8
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckenFB"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Label1(0)
    
    Label5(0).Caption = gcArtNrFiliale
    Label5(1).Caption = fnArtBezSuchen(gcArtNrFiliale)
    Label5(2).Caption = fnArtEanSuchen(gcArtNrFiliale)
    
    Label5(3).Visible = False
    Label5(4).Visible = False
    
    
    Label5(4).Caption = fnArtMBORDERSuchenMB(gcArtNrFiliale)
    If Label5(4).Caption <> "" Then
        Label5(3).Visible = True
        Label5(4).Visible = True
    End If
    
    

    
    'Aktualisieren des Bestandes aus der KissLive-Datenbank
    
    If gbKL_LIVEBESTAND = True Then
        live_bestand_abrufen gcArtNrFiliale
    End If
    
    
    BestandinFiliale gcArtNrFiliale
    LeseArtBestandinFil List1, List2
    
    FormatiereGridWKLam
    DatenforGrid gcArtNrFiliale
    anzeigenFILBESTA
    
    If Trim$(gcFilNr) = "0" Then
        LeseBestellungWKLae
    Else
    
        If NewTableSuchenDBKombi("ZBREST", gdBase) Then
            LeseZentBestellungWKLae
            
        Else
            Label4(0).Caption = "keine Info"
            Label4(0).Refresh
            
            Label4(12).Caption = "keine Info"
            Label4(12).Refresh
            
        End If
    End If
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    Resume Next
    
End Sub

Private Sub FormatiereGridWKLam()
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Rows = giAnzFil + 1
    MSFlexGrid1.Cols = 10
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Text = "Fil"
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 2000
    MSFlexGrid1.Text = "Filialname"
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 1000
    MSFlexGrid1.Text = "Bestand"
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 1200
    MSFlexGrid1.Text = "Kassen-VK"
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 900
    MSFlexGrid1.Text = "UW"
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.ColWidth(5) = 900
    MSFlexGrid1.Text = "MB"
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.ColWidth(6) = 900
    MSFlexGrid1.Text = "in BE"
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.ColWidth(7) = 1300
    MSFlexGrid1.Text = "in BE am"
    
    MSFlexGrid1.Col = 8
    MSFlexGrid1.ColWidth(8) = 900
    MSFlexGrid1.Text = "B"
    
    MSFlexGrid1.Col = 9
    MSFlexGrid1.ColWidth(9) = 900
    MSFlexGrid1.Text = "S"
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereGridWKLam"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub DatenforGrid(sArtnr As String)
    On Error GoTo LOKAL_ERROR
    
    
    If sArtnr = "" Then
        Exit Sub
    End If
    
    Dim sSQL As String
    
    loeschNEW "F" & srechnertab, gdBase
    
    sSQL = "Create Table F" & srechnertab & " ( "
    sSQL = sSQL & " Filialnr BYTE"
    sSQL = sSQL & " , Filialname Text(35)"
    sSQL = sSQL & " , Bestand Integer "
    sSQL = sSQL & " , unterwegs Integer "
    sSQL = sSQL & " , uDATE DATETIME "
    sSQL = sSQL & " , Block Text(1)"
    sSQL = sSQL & " , Sperr Text(1)"
    sSQL = sSQL & " , MB Integer "
    sSQL = sSQL & " , KVKPR1 single "
    sSQL = sSQL & " , INBEST LONG "
    sSQL = sSQL & " , INBESTAM DATETIME ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "insert into F" & srechnertab & " select filialnr,filialname "
    sSQL = sSQL & ", 0 as Bestand  "
    sSQL = sSQL & ", 0 as unterwegs  "
    sSQL = sSQL & ", null as uDATE  "
    sSQL = sSQL & ", '' as Block  "
    sSQL = sSQL & ", '' as Sperr  "
    sSQL = sSQL & ", 0 as MB  "
    sSQL = sSQL & ", 0 as KVKPR1 "
    sSQL = sSQL & ", 0 as INBEST  "
    sSQL = sSQL & ", null as INBESTAM  "
    sSQL = sSQL & " from filialen"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE F" & srechnertab & " INNER JOIN zbestand ON "
    sSQL = sSQL & " F" & srechnertab & ".filialnr = zbestand.filialnr "
    sSQL = sSQL & " Set F" & srechnertab & ".bestand = zbestand.bestand "
    sSQL = sSQL & " , F" & srechnertab & ".MB = zbestand.minbest "
    sSQL = sSQL & " , F" & srechnertab & ".KVKPR1 = zbestand.KVKPR1 "
    sSQL = sSQL & " where zbestand.artnr = " & sArtnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE F" & srechnertab & ", artikel "
    sSQL = sSQL & " SET F" & srechnertab & ".bestand = artikel.bestand "
    sSQL = sSQL & " , F" & srechnertab & ".KVKPR1 = artikel.KVKPR1 "
    sSQL = sSQL & " where F" & srechnertab & ".filialnr = " & gcFilNr
    sSQL = sSQL & " and artikel.artnr = " & sArtnr
    gdBase.Execute sSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ZUNTER", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZUNTER ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZUNTER.filiale "
        sSQL = sSQL & " Set F" & srechnertab & ".unterwegs = ZUNTER.menge "
        sSQL = sSQL & " , F" & srechnertab & ".uDATE = ZUNTER.DATUM "
        sSQL = sSQL & " where ZUNTER.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZBLOCK", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZBLOCK ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZBLOCK.filiale "
        sSQL = sSQL & " Set F" & srechnertab & ".BLOCK = 'B' "
        sSQL = sSQL & " where ZBLOCK.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZSPERR", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZSPERR ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZSPERR.filnr "
        sSQL = sSQL & " Set F" & srechnertab & ".SPERR = 'S' "
        sSQL = sSQL & " where ZSPERR.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZBREST", gdBase) Then
    
        loeschNEW "ZINBEST" & srechnertab, gdBase
    
        sSQL = "Create Table ZINBEST" & srechnertab & " ( "
        sSQL = sSQL & " ARTNR long "
        sSQL = sSQL & ", Filialnr long "
        sSQL = sSQL & ", INBEST LONG "
        sSQL = sSQL & ", INBESTAM DATETIME ) "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "insert into ZINBEST" & srechnertab & " select "
        sSQL = sSQL & " artnr "
        sSQL = sSQL & ", val(filialen) as filialnr "
        sSQL = sSQL & ", BESTVOR as INBEST  "
        sSQL = sSQL & ", BEST_datum as INBESTAM  "
        sSQL = sSQL & " from ZBREST where ARTNR = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
        
        
        loeschNEW "ZINBESTSUM" & srechnertab, gdBase
    
        sSQL = "Create Table ZINBESTSUM" & srechnertab & " ( "
        sSQL = sSQL & " ARTNR long "
        sSQL = sSQL & ", Filialnr long "
        sSQL = sSQL & ", INBESTSUM LONG "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "insert into ZINBESTSUM" & srechnertab & " select "
        sSQL = sSQL & " artnr "
        sSQL = sSQL & ", filialnr "
        sSQL = sSQL & ", sum(INBEST) as INBESTSUM  "
        sSQL = sSQL & " from ZINBEST" & srechnertab & " group by ARTNR,filialnr  "
        gdBase.Execute sSQL, dbFailOnError
        
        
        
        loeschNEW "ZINBESTMAX" & srechnertab, gdBase
    
        sSQL = "Create Table ZINBESTMAX" & srechnertab & " ( "
        sSQL = sSQL & " ARTNR long "
        sSQL = sSQL & ", Filialnr long "
        sSQL = sSQL & ", INBESTAMMAX DATETIME "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "insert into ZINBESTMAX" & srechnertab & " select "
        sSQL = sSQL & " artnr "
        sSQL = sSQL & ", filialnr "
        sSQL = sSQL & ", Max(INBESTAM) as INBESTAMMAX  "
        sSQL = sSQL & " from ZINBEST" & srechnertab & " group by ARTNR,filialnr  "
        gdBase.Execute sSQL, dbFailOnError
        
        
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZINBESTMAX" & srechnertab & " ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZINBESTMAX" & srechnertab & ".filialnr "
        sSQL = sSQL & " SET F" & srechnertab & ".INBESTAM = ZINBESTMAX" & srechnertab & ".INBESTAMMAX "
        sSQL = sSQL & " where ZINBESTMAX" & srechnertab & ".artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZINBESTSUM" & srechnertab & " ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZINBESTSUM" & srechnertab & ".filialnr "
        sSQL = sSQL & " Set F" & srechnertab & ".INBEST = ZINBESTSUM" & srechnertab & ".INBESTSUM "
        sSQL = sSQL & " where ZINBESTSUM" & srechnertab & ".artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
  
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DatenforGrid"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub anzeigenFILBESTA()
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim rsrs  As Recordset
    Dim lWert As Long
    Dim sWert As String
    Dim dWert As Double
    
    Dim sSQL As String
        
    sSQL = "Select * from F" & srechnertab & " order by filialnr"
    
    MSFlexGrid1.Redraw = False
    
    lrow = 0
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lrow = lrow + 1
            MSFlexGrid1.Rows = lrow + 1
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            
            If Not IsNull(rsrs!FILIALNR) Then
                lWert = rsrs!FILIALNR
            Else
                lWert = 0
            End If
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!Filialname) Then
                sWert = rsrs!Filialname
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = sWert
            
            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format$(dWert, "######0.00")
            
            If Not IsNull(rsrs!unterwegs) Then
                lWert = rsrs!unterwegs
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!MB) Then
                lWert = rsrs!MB
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = lWert
            
            If Not IsNull(rsrs!INBEST) Then
                lWert = rsrs!INBEST
            Else
                lWert = 0
            End If
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = lWert
            
            
            
            
            If Not IsNull(rsrs!INBESTAM) Then
                dWert = rsrs!INBESTAM
            Else
                dWert = 0
            End If
            
            If dWert > 0 Then
                sWert = Format$(dWert, "DD.MM.YYYY")
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 7
            MSFlexGrid1.Text = sWert
            
            
            
            
'            If Not IsNull(rsrs!INBESTAM) Then
'                sWert = rsrs!INBESTAM
'            Else
'                sWert = 0
'            End If
'            MSFlexGrid1.Col = 7
'            MSFlexGrid1.Text = sWert
            
            
            
            If Not IsNull(rsrs!Block) Then
                sWert = rsrs!Block
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 8
            MSFlexGrid1.Text = sWert
            
            
            If Not IsNull(rsrs!SPERR) Then
                sWert = rsrs!SPERR
            Else
                sWert = ""
            End If
            MSFlexGrid1.Col = 9
            MSFlexGrid1.Text = sWert
            
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    MSFlexGrid1.Redraw = True
    
'    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenFILBESTA"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseBestellungWKLae()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lartnr As Long
    Dim lAnz As Long
    Dim cdat As String
    
    lartnr = Val(gcArtNrFiliale)
    
    lAnz = 0
    cSQL = "Select SUM(BESTVOR) as BESTELLT from BESTREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            lAnz = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Label4(0).Caption = Trim$(Str$(lAnz))
    
    
    cdat = ""
    cSQL = "Select max(BEST_datum) as BESTELLT from BESTREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            cdat = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    If IsDate(cdat) Then
        Label4(12).Caption = DateValue(cdat)
    Else
        Label4(12).Caption = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseBestellungWKLae"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseZentBestellungWKLae()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lartnr As Long
    Dim lAnz As Long
    Dim cdat As String
    
    lartnr = Val(gcArtNrFiliale)
    
    lAnz = 0
    
    cSQL = "Select SUM(BESTVOR) as BESTELLT from ZBREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            lAnz = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Label4(0).Caption = Trim$(Str$(lAnz))
    
    
    cdat = ""
    cSQL = "Select max(BEST_datum) as BESTELLT from ZBREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            cdat = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    If IsDate(cdat) Then
        Label4(12).Caption = DateValue(cdat)
    Else
        Label4(12).Caption = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseZentBestellungWKLae"
    Fehler.gsFehlertext = "Es ist ein Fehler im Programmteil Warenbestand in den Filialen aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
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


