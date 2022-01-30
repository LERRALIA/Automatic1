VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL57 
   BackColor       =   &H00C0C0C0&
   Caption         =   "EC Lastschriften"
   ClientHeight    =   8595
   ClientLeft      =   2205
   ClientTop       =   3960
   ClientWidth     =   11880
   Icon            =   "frmWKL57.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Caption         =   "als Sepa-Datei"
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
   End
   Begin sevCommand3.Command Command4 
      Height          =   255
      Index           =   12
      Left            =   11400
      TabIndex        =   13
      Top             =   480
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "P"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdUpdate 
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "Ändern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdStandardUp 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "Standard"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtUpdatepfad 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   6975
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   6
      Top             =   3720
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
      Caption         =   "in Datei"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
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
      Caption         =   "Diskette"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   1
      Top             =   7800
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
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   7800
      Width           =   9375
   End
   Begin VB.Label lbl6 
      Caption         =   "Speicherort für die EC Lastschriftdatei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Möchten Sie die EC Lastschriften an diesem Ort  speichern, dann klicken Sie 'in Datei'!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   9375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "EC Lastschriften"
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
      TabIndex        =   4
      Top             =   120
      Width           =   7935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Label2"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmWKL57.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   9375
   End
End
Attribute VB_Name = "frmWKL57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DruckeBegleitZettelWKL57()
    On Error GoTo LOKAL_ERROR
    
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim rsrs As Recordset
    Dim cSQL As String
    Dim cTag As String
    Dim cMon As String
    Dim cJahr As String
    
    Dim aDeviceName As String
    'Dim bReturn As Boolean
    Dim cDaten As String
    Dim lAnzZeile As Long
    Dim lcount As Long
    Dim iLenZeile As Integer
    Dim cEscapeSequenz As String
    Dim cZeile As String
    ReDim cDruckZeile(1 To 1) As String
    
    
    cSQL = "Select * from BANKEN where BLZ = '" & gFirma.BLZ & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BankName) Then
            gDtaBegleit.BankName = rsrs!BankName
        Else
            gDtaBegleit.BankName = ""
        End If
    Else
        gDtaBegleit.BankName = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '//beginn
    gDtaBegleit.BegleitZettel = "DATENTRÄGERBEGLEITZETTEL"
    gDtaBegleit.BelegloserDTA = "Belegloser Datenträgeraustausch"
    gDtaBegleit.Sammel = "Sammel-Einzugsauftrag an"
    gDtaBegleit.VolNr = "DTAUS1"
    cTag = Mid(gDtaASatz.Datum, 1, 2)
    cMon = Mid(gDtaASatz.Datum, 3, 2)
    cJahr = Mid(gDtaASatz.Datum, 5, 2)
    If Val(cJahr) > 80 Then
        cJahr = "19" & cJahr
    Else
        cJahr = "20" & cJahr
    End If
    gDtaBegleit.ErstellungsDatum = cTag & "." & cMon & "." & cJahr
    gDtaBegleit.AusFuehrDatum = gDtaASatz.Erfuellung
    gDtaBegleit.AnzSatzC = Format$(Val(gDtaESatz.AnzSatz), "###,##0")
    gDtaBegleit.SummeDM = Format$(((Val(gDtaESatz.SumTotalDM)) / 100), "###,###,##0.00")
    gDtaBegleit.SummeEuro = Format$(((Val(gDtaESatz.SumTotalEuro)) / 100), "###,###,##0.00")
    gDtaBegleit.SummeKonto = Format$((Val(gDtaESatz.SumKonto)), "###,###,###,###,##0")
    gDtaBegleit.SummeBLZ = Format$((Val(gDtaESatz.SumBLZ)), "###,###,###,###,##0")
    gDtaBegleit.AbsName = gFirma.FirmaName
    gDtaBegleit.AbsBLZ = gFirma.BLZ
    gDtaBegleit.AbsKonto = gFirma.Konto
    gDtaBegleit.EmpfName = gFirma.FirmaName
    gDtaBegleit.EmpfBLZ = Space$(12 - Len(gFirma.BLZ)) & gFirma.BLZ
    gDtaBegleit.EmpfKonto = Space$(12 - Len(gFirma.Konto)) & gFirma.Konto
    gDtaBegleit.Ort = gFirma.Ort
    gDtaBegleit.Datum = Format$(Now, "DD.MM.YYYY")
    gDtaBegleit.firma = gFirma.FirmaName
    gDtaBegleit.Unterschrift = ""
    
    iLenZeile = 35
    
    '********************************************
    '*** 1.Schritt: Umschalten auf BonDrucker ***
    '********************************************
    
    setzedrucker gcBonDrucker
    '********************************************************
    '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
    '********************************************************

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz

    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    '********************************************************************
    '* Titelzeile
    '********************************************************************
    cDaten = "DATENTRÄGERBEGLEITZETTEL"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = "Sammel-Einzugsauftrag an"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = gDtaBegleit.BankName
    If Len(cDaten) > 35 Then
         cDaten = Mid$(cDaten, 1, 37)
    Else
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = "Name der DT-Austauschdatei:"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = gDtaBegleit.VolNr
    If Len(cDaten) > 35 Then
         cDaten = Mid$(cDaten, 1, 37)
    Else
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = "Erstellungsdatum:"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = gDtaBegleit.ErstellungsDatum
    If Len(cDaten) > 35 Then
         cDaten = Mid$(cDaten, 1, 37)
    Else
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = "Anzahl Lastschriften:"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = gDtaBegleit.AnzSatzC
    If Len(cDaten) > 35 Then
         cDaten = Mid$(cDaten, 1, 37)
    Else
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
   
   
    If gcWaehrung = "EUR" Then
    Else
        cDaten = "Summe in " & gcWaehrung & ": "
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = gDtaBegleit.SummeDM
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    cDaten = "Summe in EURO:"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = gDtaBegleit.SummeEuro
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Kontrollsumme KontoNr: "
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = gDtaBegleit.SummeKonto
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Kontrollsumme BLZ: "
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = gDtaBegleit.SummeBLZ
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Absender: "
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = gDtaBegleit.AbsName
    If Len(cDaten) > 35 Then
         cDaten = Mid$(cDaten, 1, 37)
    Else
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Bankleitzahl: " & gDtaBegleit.AbsBLZ
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Kontonummer: " & gDtaBegleit.AbsKonto
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '//Leerzeilen
    For lcount = 1 To 10
        If lcount = 10 Then
            cEscapeSequenz = "." & vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
    cDaten = "_______________________________"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cDaten = "(Ort, Datum, Unterschrift) "
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '//Leerzeilen
    For lcount = 1 To 9
        If lcount = 9 Then
            cEscapeSequenz = "." & vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
'    aDeviceName = gcBonDrucker
'    bReturn = SetDefaultPrinter(aDeviceName)
'    If Not bReturn Then
'        MsgBox "Drucker für Ausdruck 'Anlage zum Begleitzettel' ist nicht bereit!", vbCritical, "STOP!"
'        Exit Sub
'    End If
    
    '//drucken
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    
    '//schneiden
    If gbAPI = True Then
        aDeviceName = gcBonDrucker
        'aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    '//end

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeBegleitzettelWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub DruckeBegleitZettelAnlageWKL57(cKonto() As String, cBLZ() As String, cBetrag() As String)
    On Error GoTo LOKAL_ERROR
    
    Dim aDeviceName As String
    Dim bReturn As Boolean
    Dim cDaten As String
    Dim lAnzZeile As Long
    Dim lcount As Long
    Dim iLenZeile As Integer
    Dim cEscapeSequenz As String
    Dim cZeile As String
    ReDim cDruckZeile(1 To 1) As String
    
    iLenZeile = 35
    
    '********************************************
    '*** 1.Schritt: Umschalten auf BonDrucker ***
    '********************************************
    
    setzedrucker gcBonDrucker
    '********************************************************
    '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
    '********************************************************

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz

    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    '********************************************************************
    '* Titelzeile
    '********************************************************************
    cDaten = "Anlage zum DT-Begleitzettel"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '********************************************************************
    '* Datum
    '********************************************************************
    
    cDaten = "vom " & Format$(Fix(Now), "DD.MM.YYYY")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* doppelter Trennstrich für Überschriften
    '********************************************************************
    
    cDaten = String$(35, "=")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* Überschrift
    '********************************************************************
    
    cDaten = "Konto         Bank          Betrag"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* doppelter Trennstrich für Überschriften
    '********************************************************************
    
    cDaten = String$(35, "=")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* Daten einlesen
    '********************************************************************
    
    For lcount = LBound(cKonto) To UBound(cKonto)
        cZeile = cKonto(lcount) & "  " & cBLZ(lcount) & "  " & cBetrag(lcount)
        
        cDaten = cZeile
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = String$(35, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz

    Next lcount
    
    '********************************************************************
    '* Leerzeile drucken
    '********************************************************************
    
    cDaten = Space$(35)
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* doppelter Trennstrich
    '********************************************************************
    
    cDaten = String$(35, "=")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '********************************************************************
    '* Datum
    '********************************************************************
    
    cDaten = Format$(Fix(Now), "DD.MM.YYYY") & "  " & Format$(Now, "HH:MM") & " " & gcKasNum
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* doppelter Trennstrich
    '********************************************************************
    
    cDaten = String$(35, "=")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '********************************************************************
    '* Drucker schalten
    '********************************************************************
    
    aDeviceName = gcBonDrucker
    bReturn = SetDefaultPrinter(aDeviceName)
    If Not bReturn Then
        MsgBox "Drucker für Ausdruck 'Anlage zum Begleitzettel' ist nicht bereit!", vbCritical, "STOP!"
        Exit Sub
    End If
    
    
    '//Leerzeilen
    For lcount = 1 To 9
        If lcount = 9 Then
            cEscapeSequenz = "." & vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    '//schneiden
    If gbAPI Then
        aDeviceName = gcBonDrucker
        'aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
   

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeBegleitzettelAnlageWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub HoleDtaASatzWKL57()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    ctmp = gFirma.FirmaName
    ctmp = UCase$(ctmp)
    KonvertAnsiAscii ctmp
    
    gDtaASatz.SatzLen = "0128"
    gDtaASatz.SatzArt = "A"
    gDtaASatz.Hinweis = "LK"
    gDtaASatz.BLZ_Empf = gFirma.BLZ
    gDtaASatz.Filler1 = String$(8, "0")
    If Len(ctmp) <= 27 Then
        gDtaASatz.Absender = ctmp
    Else
        gDtaASatz.Absender = Left(ctmp, 27)
    End If
    gDtaASatz.Datum = Format$(Now, "DDMMYY")
    gDtaASatz.Filler2 = Space$(4)
    gDtaASatz.KontoEmpf = String$(10 - Len(gFirma.Konto), "0") & gFirma.Konto
    gDtaASatz.RefNr = String$(10, "0")
    'gDtaASatz.Reserve1 = Space$(15)
    gDtaASatz.Reserve1 = Space$(47)

    gDtaASatz.Erfuellung = Format$(Now, "DDMMYYYY")
    'gDtaASatz.Reserve2 = Space$(24)
    
    If gcWaehrung = "EUR" Then
        gDtaASatz.WaeCode = "1"
    Else
        gDtaASatz.WaeCode = Space$(1)
    End If

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDtaASatzWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub HoleDtaESatzWKL57(lAnzCSatz As Long, dSumTotalDM As Double, dSumKonto As Double, dSumBLZ As Double, dSumTotalEuro As Double)
    On Error GoTo LOKAL_ERROR
    
    gDtaESatz.SatzLen = "0128"
    gDtaESatz.SatzArt = "E"
    gDtaESatz.Filler1 = Space$(5)
    gDtaESatz.AnzSatz = Format$(lAnzCSatz, "0000000")
    
    If gcWaehrung = "EUR" Then
        gDtaESatz.SumTotalDM = String$(13, "0")
    Else
        gDtaESatz.SumTotalDM = Format$(dSumTotalDM, "0000000000000")
    End If
    gDtaESatz.SumKonto = Format$(dSumKonto, "00000000000000000")
    gDtaESatz.SumBLZ = Format$(dSumBLZ, "00000000000000000")
    gDtaESatz.SumTotalEuro = Format$(dSumTotalEuro, "0000000000000")
    gDtaESatz.Filler2 = Space$(51)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDtaESatzWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub LeereDTASaetze()
    On Error GoTo LOKAL_ERROR
    
    gDtaASatz.SatzLen = "0128"
    gDtaASatz.SatzArt = "A"
    gDtaASatz.Hinweis = "LK"
    gDtaASatz.BLZ_Empf = String$(8, "0")
    gDtaASatz.Filler1 = String$(8, "0")
    gDtaASatz.Absender = Space$(27)
    gDtaASatz.Datum = String$(6, "0")
    gDtaASatz.Filler2 = Space$(4)
    gDtaASatz.KontoEmpf = String$(10, "0")
    gDtaASatz.RefNr = Space$(10)
    gDtaASatz.Reserve1 = Space$(15)
    gDtaASatz.Erfuellung = String$(8, "0")
    gDtaASatz.Reserve2 = Space$(24)

    gDtaESatz.SatzLen = "0128"
    gDtaESatz.SatzArt = "E"
    gDtaESatz.Filler1 = Space$(5)
    gDtaESatz.AnzSatz = String$(7, "0")
    gDtaESatz.SumTotalDM = String$(13, "0")
    gDtaESatz.SumKonto = String$(17, "0")
    gDtaESatz.SumBLZ = String$(17, "0")
    gDtaESatz.SumTotalEuro = String$(13, "0")
    gDtaESatz.Filler2 = Space$(51)


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDTASaetzeWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub LeereDTASaetzeC()
    On Error GoTo LOKAL_ERROR
    
    gDtaCSatz.SatzLen = "0187"
    gDtaCSatz.SatzArt = "C"
    gDtaCSatz.KdBLZ1 = String$(8, "0")
    gDtaCSatz.KdBLZ2 = String$(8, "0")
    gDtaCSatz.Konto = String$(10, "0")
    gDtaCSatz.KdNr = String$(13, "0")
    gDtaCSatz.TextKey = Space$(2)
    gDtaCSatz.TextKeyAdd = Space$(3)
    gDtaCSatz.Filler1 = Space$(1)
    gDtaCSatz.BetragDM = String$(11, "0")
    gDtaCSatz.EmpfBLZ = String$(8, "0")
    gDtaCSatz.EmpfKonto = String$(10, "0")
    gDtaCSatz.BetragEuro = String$(11, "0")
    gDtaCSatz.Filler2 = Space$(3)
    gDtaCSatz.Empfaenger = Space$(27)
    gDtaCSatz.Filler3 = Space$(8)
    gDtaCSatz.AuftragName = Space$(27)
    gDtaCSatz.Zweck = Space$(27)
    gDtaCSatz.WaeCode = Space$(1)
    gDtaCSatz.Filler4 = Space$(2)
    gDtaCSatz.AnzErweit = "00"
    gDtaCSatz.KzErwTeil1 = "03"
    gDtaCSatz.Wahltext1 = Space$(27)
    gDtaCSatz.KzErwTeil2 = Space$(2)
    gDtaCSatz.Wahltext2 = Space$(27)
    gDtaCSatz.Filler5 = Space$(11)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDTASaetzeCWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SchreibeDBF2DTADiskWKL57()
    On Error GoTo LOKAL_ERROR
    
    Dim cZweiMal As String
    Dim lRet As Long
    Dim bDiskOk As Boolean
    Dim iFileNr As Integer
    Dim lPos As Long
    Dim ctmp As String
    
    Dim lAnzCSatz As Long
    Dim dSumTotalDM As Double
    Dim dSumTotalEuro As Double
    Dim dSumKonto As Double
    Dim dSumBLZ As Double
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim dWert As Double

    ReDim cKonto(1 To 1) As String
    ReDim cBLZ(1 To 1) As String
    ReDim cBetrag(1 To 1) As String
    Dim lAnzSatz As Long
    Dim cQuelle As String
    Dim cZiel   As String
    Dim lfail As Long
    
    
    bDiskOk = True

    File1.Path = "A:"

    If Not bDiskOk Then
        Screen.MousePointer = 0
        MsgBox "Bitte eine Diskette in Laufwerk A: einlegen!", vbCritical, "STOP!"
        Exit Sub
    End If

    If File1.ListCount > 0 Then
        'Disk ist nicht leer!
        Screen.MousePointer = 0
        lRet = MsgBox("Diskette ist nicht leer! Formatieren?", vbYesNo + vbQuestion, "FORMATIEREN")
        If lRet <> vbYes Then
            Exit Sub
        Else
            SyncShell "Format.com a:", 0, False, False
'            lRet = Shell("Format.com a:", 3)
            Screen.MousePointer = 0
            MsgBox "Diskette formatiert!", vbInformation, "Winkiss Hinweis:"
            File1.Refresh
        End If
    End If

    LeereDTASaetze
    LeereDTASaetzeC
    
    sicherdta

    HoleDtaASatzWKL57

    lAnzCSatz = 0
    dSumTotalDM = 0
    dSumTotalEuro = 0
    dSumKonto = 0
    dSumBLZ = 0

    iFileNr = FreeFile
    Open "A:\DTAUS1" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, gDtaASatz
    lPos = lPos + Len(gDtaASatz)

    cSQL = "Select * from DTA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzSatz = 0
        Do While Not rsrs.EOF
            lAnzSatz = lAnzSatz + 1
            ReDim Preserve cKonto(1 To lAnzSatz) As String
            ReDim Preserve cBLZ(1 To lAnzSatz) As String
            ReDim Preserve cBetrag(1 To lAnzSatz) As String

            LeereDTASaetzeC

            lAnzCSatz = lAnzCSatz + 1

            gDtaCSatz.SatzLen = "0216"                  'Feld C1
            gDtaCSatz.SatzArt = "C"                     'Feld C2
            gDtaCSatz.KdBLZ1 = gFirma.BLZ               'Feld C3

            If Not IsNull(rsrs!BLZ) Then
                gDtaCSatz.KdBLZ2 = rsrs!BLZ             'Feld C4
                dSumBLZ = dSumBLZ + Val(rsrs!BLZ)
                cBLZ(lAnzSatz) = rsrs!BLZ
                cBLZ(lAnzSatz) = Space$(10 - Len(cBLZ(lAnzSatz))) & cBLZ(lAnzSatz)
            End If
            If Not IsNull(rsrs!EKONTO) Then
                ctmp = rsrs!EKONTO
                ctmp = Trim$(ctmp)
                cKonto(lAnzSatz) = ctmp
                cKonto(lAnzSatz) = Space$(10 - Len(cKonto(lAnzSatz))) & cKonto(lAnzSatz)
                ctmp = String$(10 - Len(ctmp), "0") & ctmp
                gDtaCSatz.Konto = ctmp                  'Feld C5
                dSumKonto = dSumKonto + Val(rsrs!EKONTO)
            Else
                ctmp = "0000000000"
                gDtaCSatz.Konto = ctmp
            End If

            gDtaCSatz.KdNr = String$(13, "0")           'Feld C6

            If Not IsNull(rsrs!TextA) Then
                If rsrs!TextA = "5" Then
                    gDtaCSatz.TextKey = "05"            'Feld C7a
                    gDtaCSatz.TextKeyAdd = "000"        'Feld C7b
                ElseIf rsrs!TextA = "4" Then
                    gDtaCSatz.TextKey = "04"            'Feld C7a
                    gDtaCSatz.TextKeyAdd = "000"        'Feld C7b
                Else
                    gDtaCSatz.TextKey = "51"            'Feld C7a
                    gDtaCSatz.TextKeyAdd = "000"        'Feld C7b
                End If
            End If


            gDtaCSatz.Filler1 = Space$(1)                'Feld C8


            If gcWaehrung = "EUR" Then
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)

                dWert = dWert * 100
                dSumTotalDM = dSumTotalDM + dWert
                gDtaCSatz.BetragDM = String$(11, "0")   'Feld C9
            Else
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)

                dWert = dWert * 100
                dSumTotalDM = dSumTotalDM + dWert
                gDtaCSatz.BetragDM = Format$(dWert, "00000000000")  'Feld C9
            End If

            gDtaCSatz.EmpfBLZ = String$(8 - Len(gFirma.BLZ), "0") & gFirma.BLZ            'Feld C10
            gDtaCSatz.EmpfKonto = String$(10 - Len(gFirma.Konto), "0") & gFirma.Konto     'Feld C11

            If gcWaehrung = "EUR" Then
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)
                dWert = dWert * 100
                dSumTotalEuro = dSumTotalEuro + dWert
                gDtaCSatz.BetragEuro = Format$(dWert, "00000000000")  'Feld C12
            Else
                gDtaCSatz.BetragEuro = String$(11, "0")   'Feld C12
            End If

            gDtaCSatz.Filler2 = Space$(3)               'Feld C13

            If Not IsNull(rsrs!Empfaenger) Then         'Feld C14a
                ctmp = rsrs!Empfaenger
                ctmp = UCase$(Trim$(ctmp))
                ctmp = ctmp & Space$(27 - Len(ctmp))
                KonvertAnsiAscii ctmp
                gDtaCSatz.Empfaenger = ctmp
            Else
                gDtaCSatz.Empfaenger = Space$(27)
            End If

            gDtaCSatz.Filler3 = Space$(8)               'Feld C14b

            ctmp = gFirma.FirmaName
            ctmp = UCase$(Trim$(ctmp))
            KonvertAnsiAscii ctmp
            If Len(gFirma.FirmaName) > 27 Then          'Feld C15
                gDtaCSatz.AuftragName = Left(ctmp, 27)
            Else
                gDtaCSatz.AuftragName = ctmp & Space$(27 - Len(ctmp))
            End If

            If Not IsNull(rsrs!zweck1) Then
                ctmp = "Danke "
                ctmp = ctmp & rsrs!zweck1
                If Not IsNull(rsrs!FILIALE) Then
                    ctmp = ctmp & "/" & rsrs!FILIALE
                End If
                
                If Not IsNull(rsrs!Datum) Then
                    ctmp = ctmp & " " & rsrs!Datum
                End If
'                ctmp = UCase$(Trim$(ctmp))
                ctmp = ctmp & Space$(27 - Len(ctmp))
                KonvertAnsiAscii ctmp
                gDtaCSatz.Zweck = ctmp
            End If

            '//Aenderung
            If gcWaehrung = "DEM" Or gcWaehrung = "ATS" Or gcWaehrung = "NLG" Or gcWaehrung = "CHF" Then
                gDtaCSatz.WaeCode = Space$(1)
            ElseIf gcWaehrung = "EUR" Then
                gDtaCSatz.WaeCode = "1"
            Else
            End If

            gDtaCSatz.Filler4 = Space$(2)
            gDtaCSatz.AnzErweit = "01"
            ctmp = gFirma.strasse
            ctmp = Trim$(UCase$(ctmp))
            ctmp = ctmp & Space$(27 - Len(ctmp))
            KonvertAnsiAscii ctmp
            gDtaCSatz.Wahltext1 = ctmp

            Put #iFileNr, lPos, gDtaCSatz

            rsrs.MoveNext

            lPos = lPos + Len(gDtaCSatz)
        Loop
    Else

    End If
    rsrs.Close: Set rsrs = Nothing

    HoleDtaESatzWKL57 lAnzCSatz, dSumTotalDM, dSumKonto, dSumBLZ, dSumTotalEuro
    
    Put #iFileNr, lPos, gDtaESatz
    Close iFileNr
    
    '//Kopieren nach C:\Eigene Dateien
    Open "A:\DTAUS1" For Binary As #iFileNr
    Get #iFileNr, lPos, gDtaESatz
    
    
    cQuelle = "A:\DTAUS1"
    
    Dim cPfad As String
    Dim lWert As Long
    Dim cdatei As String
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "DTA" & ctmp & Format$(TimeValue(Now), "HH:MM")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DTA\"
    cZiel = cPfad & cdatei
    
    
    lRet = CopyFile(cQuelle, cZiel, lfail)
    '//End Kopieren
    
    cZweiMal = "ok"
    
    KompressDTAWKL57
    
    MsgBox "DTA-Diskette erstellt!" & vbCrLf & vbCrLf & "Begleitzettel wird gedruckt...", vbInformation, "FERTIG!"
    
ZweiteBon:
    DruckeBegleitZettelWKL57
    
    DruckeBegleitZettelAnlageWKL57 cKonto(), cBLZ(), cBetrag()
    
    If cZweiMal = "ok" Then
        cZweiMal = ""
        GoTo ZweiteBon
    Else
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 68 Then
        bDiskOk = False
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeDBF2DTADiskWKL57"
        Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub SchreibeDTAinPfad()
    On Error GoTo LOKAL_ERROR
    
    Dim cZweiMal As String
    Dim lRet As Long
    Dim iRet As Integer
    Dim iFileNr As Integer
    Dim lPos As Long
    Dim ctmp As String
    
    Dim lAnzCSatz As Long
    Dim dSumTotalDM As Double
    Dim dSumTotalEuro As Double
    Dim dSumKonto As Double
    Dim dSumBLZ As Double
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim dWert As Double

    ReDim cKonto(1 To 1) As String
    ReDim cBLZ(1 To 1) As String
    ReDim cBetrag(1 To 1) As String
    Dim lAnzSatz As Long
    Dim cQuelle As String
    Dim cZiel   As String
    Dim lfail As Long

    LeereDTASaetze
    LeereDTASaetzeC
    
    sicherdta

    HoleDtaASatzWKL57

    lAnzCSatz = 0
    dSumTotalDM = 0
    dSumTotalEuro = 0
    dSumKonto = 0
    dSumBLZ = 0
    
    Kill gsDTAPfad & "\*.*"
    
    iFileNr = FreeFile
    Open gsDTAPfad & "\DTAUS1" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, gDtaASatz
    lPos = lPos + Len(gDtaASatz)

    cSQL = "Select * from DTA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzSatz = 0
        Do While Not rsrs.EOF
            lAnzSatz = lAnzSatz + 1
            ReDim Preserve cKonto(1 To lAnzSatz) As String
            ReDim Preserve cBLZ(1 To lAnzSatz) As String
            ReDim Preserve cBetrag(1 To lAnzSatz) As String

            LeereDTASaetzeC

            lAnzCSatz = lAnzCSatz + 1

            gDtaCSatz.SatzLen = "0216"                  'Feld C1
            gDtaCSatz.SatzArt = "C"                     'Feld C2
            gDtaCSatz.KdBLZ1 = gFirma.BLZ               'Feld C3

            If Not IsNull(rsrs!BLZ) Then
                gDtaCSatz.KdBLZ2 = rsrs!BLZ             'Feld C4
                dSumBLZ = dSumBLZ + Val(rsrs!BLZ)
                cBLZ(lAnzSatz) = rsrs!BLZ
                cBLZ(lAnzSatz) = Space$(10 - Len(cBLZ(lAnzSatz))) & cBLZ(lAnzSatz)
            End If
            If Not IsNull(rsrs!EKONTO) Then
                ctmp = rsrs!EKONTO
                ctmp = Trim$(ctmp)
                cKonto(lAnzSatz) = ctmp
                cKonto(lAnzSatz) = Space$(10 - Len(cKonto(lAnzSatz))) & cKonto(lAnzSatz)
                ctmp = String$(10 - Len(ctmp), "0") & ctmp
                gDtaCSatz.Konto = ctmp                  'Feld C5
                dSumKonto = dSumKonto + Val(rsrs!EKONTO)
            Else
                ctmp = "0000000000"
                gDtaCSatz.Konto = ctmp
            End If

            gDtaCSatz.KdNr = String$(13, "0")           'Feld C6

            If Not IsNull(rsrs!TextA) Then
                If rsrs!TextA = "5" Then
                    gDtaCSatz.TextKey = "05"            'Feld C7a
                    If Check1.Value = vbChecked Then
                        gDtaCSatz.TextKeyAdd = "019"    'Feld C7b
                    Else
                        gDtaCSatz.TextKeyAdd = "000"    'Feld C7b
                    End If
                ElseIf rsrs!TextA = "4" Then
                    gDtaCSatz.TextKey = "04"            'Feld C7a
                    gDtaCSatz.TextKeyAdd = "000"        'Feld C7b
                Else
                    gDtaCSatz.TextKey = "51"            'Feld C7a
                    gDtaCSatz.TextKeyAdd = "000"        'Feld C7b
                End If
            End If


            gDtaCSatz.Filler1 = Space$(1)                'Feld C8


            If gcWaehrung = "EUR" Then
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)

                dWert = dWert * 100
                dSumTotalDM = dSumTotalDM + dWert
                gDtaCSatz.BetragDM = String$(11, "0")   'Feld C9
            Else
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)

                dWert = dWert * 100
                dSumTotalDM = dSumTotalDM + dWert
                gDtaCSatz.BetragDM = Format$(dWert, "00000000000")  'Feld C9
            End If

            gDtaCSatz.EmpfBLZ = String$(8 - Len(gFirma.BLZ), "0") & gFirma.BLZ            'Feld C10
            gDtaCSatz.EmpfKonto = String$(10 - Len(gFirma.Konto), "0") & gFirma.Konto     'Feld C11

            If gcWaehrung = "EUR" Then
                If Not IsNull(rsrs!Betrag) Then
                    dWert = rsrs!Betrag
                Else
                    dWert = 0
                End If
                cBetrag(lAnzSatz) = Format$(dWert, "######0.00")
                cBetrag(lAnzSatz) = Space$(10 - Len(cBetrag(lAnzSatz))) & cBetrag(lAnzSatz)
                dWert = dWert * 100
                dSumTotalEuro = dSumTotalEuro + dWert
                gDtaCSatz.BetragEuro = Format$(dWert, "00000000000")  'Feld C12
            Else
                gDtaCSatz.BetragEuro = String$(11, "0")   'Feld C12
            End If

            gDtaCSatz.Filler2 = Space$(3)               'Feld C13

            If Not IsNull(rsrs!Empfaenger) Then         'Feld C14a
                ctmp = rsrs!Empfaenger
                ctmp = UCase$(Trim$(ctmp))
                ctmp = ctmp & Space$(27 - Len(ctmp))
                KonvertAnsiAscii ctmp
                gDtaCSatz.Empfaenger = ctmp
            Else
                gDtaCSatz.Empfaenger = Space$(27)
            End If

            gDtaCSatz.Filler3 = Space$(8)               'Feld C14b

            ctmp = gFirma.FirmaName
            ctmp = UCase$(Trim$(ctmp))
            KonvertAnsiAscii ctmp
            
            If Len(gFirma.FirmaName) > 27 Then          'Feld C15
                gDtaCSatz.AuftragName = Left(ctmp, 27)
            Else
                gDtaCSatz.AuftragName = ctmp & Space$(27 - Len(ctmp))
            End If

            If Not IsNull(rsrs!zweck1) Then
                ctmp = "Danke "
                ctmp = ctmp & rsrs!zweck1
                If Not IsNull(rsrs!FILIALE) Then
                    ctmp = ctmp & "/" & rsrs!FILIALE
                End If
                
                If Not IsNull(rsrs!Datum) Then
                    ctmp = ctmp & " " & rsrs!Datum
                End If
'                ctmp = UCase$(Trim$(ctmp))
                ctmp = ctmp & Space$(27 - Len(ctmp))
                KonvertAnsiAscii ctmp
                gDtaCSatz.Zweck = ctmp
            End If

            '//Aenderung
            If gcWaehrung = "DEM" Or gcWaehrung = "ATS" Or gcWaehrung = "NLG" Or gcWaehrung = "CHF" Then
                gDtaCSatz.WaeCode = Space$(1)
            ElseIf gcWaehrung = "EUR" Then
                gDtaCSatz.WaeCode = "1"
            Else
            End If

            gDtaCSatz.Filler4 = Space$(2)
            gDtaCSatz.AnzErweit = "01"
            ctmp = gFirma.strasse
            ctmp = Trim$(UCase$(ctmp))
            ctmp = ctmp & Space$(27 - Len(ctmp))
            KonvertAnsiAscii ctmp
            gDtaCSatz.Wahltext1 = ctmp

            Put #iFileNr, lPos, gDtaCSatz

            rsrs.MoveNext

            lPos = lPos + Len(gDtaCSatz)
        Loop
    Else

    End If
    rsrs.Close: Set rsrs = Nothing

    HoleDtaESatzWKL57 lAnzCSatz, dSumTotalDM, dSumKonto, dSumBLZ, dSumTotalEuro
    
    Put #iFileNr, lPos, gDtaESatz
    Close iFileNr
    
    '//Kopieren nach Dta
    Open gsDTAPfad & "\DTAUS1" For Binary As #iFileNr
    Get #iFileNr, lPos, gDtaESatz
    
    
    cQuelle = gsDTAPfad & "\DTAUS1"
    
    Dim cPfad As String
    Dim lWert As Long
    Dim cdatei As String
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "DTA" & ctmp & Format$(TimeValue(Now), "HH:MM")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DTA\"
    cZiel = cPfad & cdatei
    
    Close iFileNr
    
    
    lRet = CopyFile(cQuelle, cZiel, lfail)
    '//End Kopieren
    
    cZweiMal = "ok"
    
    KompressDTAWKL57
    
    iRet = MsgBox("EC Lastenschriften sind gespeichert unter:" _
    & vbCrLf & vbCrLf & gsDTAPfad & vbCrLf & vbCrLf _
    & "Möchten Sie den Begleitzettel erstellen?", vbQuestion + vbYesNo, "Winkiss Hinweis:")
    
    If iRet = vbYes Then
ZweiteBon:
    
        DruckeBegleitZettelWKL57
        DruckeBegleitZettelAnlageWKL57 cKonto(), cBLZ(), cBetrag()
        
        If cZweiMal = "ok" Then
            cZweiMal = ""
            GoTo ZweiteBon
        Else
        
        End If
    Else
        schreibeFilProt "kein Begleitzettel drucken", "BegleitProt"
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeDTAinPfad"
        Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub cmdStandardUp_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    txtUpdatepfad.Text = cPfad & "DTAHEUTE"
    gsDTAPfad = cPfad & "DTAHEUTE"
    
    speicherpfad

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdStandardUp_Click"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo LOKAL_ERROR

    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    sTitle = "Speichern des EC Lastschriftenpfades"
    sFilter = "alle Dateien (*.*)| *.*"
    sOldpfad = txtUpdatepfad.Text

    gsDTAPfad = pfadaendern(sTitle, sFilter, sOldpfad)
    
    
    
    txtUpdatepfad.Text = gsDTAPfad
    speicherpfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdUpdate_Click"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim iFileNr As Integer
    
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
        
        
        
            If Trim(txtUpdatepfad.Text) = "" Then
                cmdStandardUp_Click
            End If
            
            
            
        
            txtUpdatepfad.Text = Trim(txtUpdatepfad.Text)
            If Right(txtUpdatepfad.Text, 1) = "\" Then
                txtUpdatepfad.Text = Left(txtUpdatepfad.Text, Len(txtUpdatepfad.Text) - 1)
            End If
            
            iFileNr = FreeFile
            Open txtUpdatepfad.Text & "\SCHNICK.TXT" For Binary As #iFileNr
            Close iFileNr
            
            Kill txtUpdatepfad.Text & "\SCHNICK.TXT"
            
            
            gsDTAPfad = txtUpdatepfad.Text
            speicherpfad
            
            Command1(0).Enabled = False
        
            SchreibeDTAinPfad
            Unload frmWKL57
            
        Case Is = 2
            
            SchreibeDBF2DTADiskWKL57
            Unload frmWKL57
            
        Case Is = 1
            Unload frmWKL57
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 76 Then
        Screen.MousePointer = 0
        MsgBox "Die Pfadangabe ist nicht korrekt." & vbCrLf & "Bitte geben Sie den Pfad nochmals ein!", vbInformation, "Winkiss Hinweis:"
        txtUpdatepfad.SetFocus
    Else

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim cPfad As String

cPfad = gcDBPfad
If Right(cPfad, 1) <> "\" Then
    cPfad = cPfad & "\"
End If

zeigeHilfe "LPROTOK", "BegleitProt.txt", cPfad
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Skalieren Me, True, True: Schrift Me: Modul6.Log Me
    Farbform Me, lblUeberschrift
    
    txtUpdatepfad.Text = gsDTAPfad
    
    If gbECTOZ Then
        Command1(2).Enabled = False
        Command1(0).Enabled = False
    Else
        Command1(2).Enabled = True
        Command1(0).Enabled = True
    End If
    
    If gFirma.BIC <> "" Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    
    Label2.Caption = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Setze Standarddrucker, bitte warten...", Label4
    
    LogtoEnd Me
    setzedrucker gcListenDrucker
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub txtUpdatepfad_lostfocus()
    txtUpdatepfad.BackColor = vbWhite
End Sub

Private Sub txtUpdatepfad_GotFocus()
    
    txtUpdatepfad.BackColor = glSelBack1

End Sub
