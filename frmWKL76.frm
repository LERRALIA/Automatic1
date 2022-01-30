VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL76 
   Caption         =   "Artikellistengenerator"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL76.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   7095
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3015
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
         Caption         =   "Preisliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bezeichnung"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linie, Bezeichnung"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   3015
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
      Caption         =   "abgewogene Artikel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   9000
      TabIndex        =   7
      Top             =   960
      Width           =   2655
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
      Caption         =   "alle Artikel in Excel exportieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
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
      Caption         =   "Konditionenliste"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
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
      Caption         =   "Linienliste"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   2
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
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
      Caption         =   "Excel Export"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
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
      Caption         =   "Lagerplatzliste"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   9255
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikellistengenerator"
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
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL76"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL76
        Case 1 'excel Export
            ExcelExport
        Case 2 'Lagerplatzliste
            LagerPlatzListe
        Case 3 'Linienliste
            Linienliste
        Case 4 'Preisliste
            Preisliste
        Case 5 'Konditionenliste
            Konditionenliste
        Case 6 'excel Export alle Artikel
            ExcelExportalleArt
        Case 7 'Gewichte
            Gewichteliste
    End Select
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub ExcelExportalleArt()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim cDatname    As String
    
    cDatname = "alleArtikel" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    frmWKL69.Show 1
        
        
    Select Case Month(DateValue(Now))
        Case 1
            If gsZeitPass <> "Schlümpfe" Then
                Exit Sub
            End If
        Case 2
            If gsZeitPass <> "Single" Then
                Exit Sub
            End If
        Case 3
            If gsZeitPass <> "Ölschock" Then
                Exit Sub
            End If
        Case 4
            If gsZeitPass <> "Zweierkiste" Then
                Exit Sub
            End If
        Case 5
            If gsZeitPass <> "Mitte" Then
                Exit Sub
            End If
        Case 6
            If gsZeitPass <> "Wende" Then
                Exit Sub
            End If
        Case 7
            If gsZeitPass <> "Havarie" Then
                Exit Sub
            End If
        Case 8
            If gsZeitPass <> "Waldsterben" Then
                Exit Sub
            End If
        Case 9
            If gsZeitPass <> "Molkepulver" Then
                Exit Sub
            End If
        Case 10
            If gsZeitPass <> "Tiefflug" Then
                Exit Sub
            End If
        Case 11
            If gsZeitPass <> "Realo" Then
                Exit Sub
            End If
        Case 12
            If gsZeitPass <> "Eurogeld" Then
                Exit Sub
            End If
    End Select
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select Artnr "
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , Libesnr"
    sSQL = sSQL & " , EAN"
    sSQL = sSQL & " , EAN2"
    sSQL = sSQL & " , EAN3"
    sSQL = sSQL & " , BESTAND"
    sSQL = sSQL & " , LEKPR"
    sSQL = sSQL & " , KVKPR1"
    sSQL = sSQL & " , VKPR"
    sSQL = sSQL & " , MWST"
    sSQL = sSQL & " , LINR"
    sSQL = sSQL & " , gefuehrt"
    sSQL = sSQL & " , RKZ"
    sSQL = sSQL & " , '' as Artikelstatus "
    sSQL = sSQL & " into ArtExcA from ARTIKEL "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA"
    sSQL = sSQL & " set EAN = '' "
    sSQL = sSQL & " , EAN2 = '' "
    sSQL = sSQL & " , EAN3 = '' "
    gdBase.Execute sSQL, dbFailOnError

    cdatei = cPfad1 & "BOX\" & cDatname
    cPfad = cPfad1 & "BOX"
    
    sSQL = "Select * into ArtExcA IN '" & cdatei & "' 'Excel 8.0;' from ArtExcA "
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
    loeschNEW "ArtExcA", gdBase
    
    Screen.MousePointer = 0
    

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExportalleArt"
        Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExcelExport()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim i           As Integer
    Dim cDatname    As String
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    cDatname = "Artikel" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"

    If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
    
    
        If SpalteInTabellegefundenNEW("TOPZ" & srechnertab, "Inhalt", gdBase) = False Then
            SpalteAnfuegenNEW "TOPZ" & srechnertab, "Inhalt", "double", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("TOPZ" & srechnertab, "Inhaltbez", gdBase) = False Then
            SpalteAnfuegenNEW "TOPZ" & srechnertab, "Inhaltbez", "Text(3)", gdBase
        End If
        
        
    
        sSQL = "Update TOPZ" & srechnertab & " inner join Artikel on TOPZ" & srechnertab & ".Artnr = Artikel.Artnr "
        sSQL = sSQL & "  Set  TOPZ" & srechnertab & ".INHALT = Artikel.INHALT "
        sSQL = sSQL & ", TOPZ" & srechnertab & ".INHALTBEZ = Artikel.INHALTBEZ "
        gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
        
    
    
    
    
        loeschNEW "ArtExc", gdBase
        
        gsZSpalte = ""
        gstab = "ARTEX"
        frmWKL36.Show 1
        
        'dannach Tablay auswerten
        Tabcheck "ARTEX"
        FormatGridOverTablay "ARTEX"
        
        If byAnzahlSpalten > 0 Then
            sSQL = "Select " & sSpaltenbez(0) & " "
            
            If byAnzahlSpalten > 1 Then
                For i = 1 To byAnzahlSpalten - 1
                    sSQL = sSQL & " , " & sSpaltenbez(i) & " "
                Next i
            End If
        Else
            Exit Sub
        End If
        
        sSQL = sSQL & " into ArtExc from TOPZ" & srechnertab
        gdBase.Execute sSQL, dbFailOnError
    
        cdatei = cPfad1 & "BOX\" & cDatname
        cPfad = cPfad1 & "BOX"
        
        sSQL = "Select * into ArtExc IN '" & cdatei & "' 'Excel 8.0;' from ArtExc "
        gdBase.Execute sSQL, dbFailOnError

    
        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
        loeschNEW "ArtExc", gdBase
        
        If SpalteInTabellegefundenNEW("TOPZ" & srechnertab, "Inhalt", gdBase) = True Then
            sSQL = " Alter table TOPZ" & srechnertab & " drop Inhalt  "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("TOPZ" & srechnertab, "Inhaltbez", gdBase) = True Then
            sSQL = " Alter table TOPZ" & srechnertab & " drop Inhaltbez  "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExport"
        Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub LagerPlatzListe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cLief As String

    If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
    
        loeschNEW "ArtExc", gdBase
        
        sSQL = "Select Lagerp,Artnr,Bezeich,Libesnr into ArtExc from TOPZ" & srechnertab
        sSQL = sSQL & " order by Lagerp "
        gdBase.Execute sSQL, dbFailOnError
        
        cLief = ermLief
        loeschNEW "LIEFExc", gdBase
        CreateTable "LIEFEXC", gdBase

        If cLief <> "0" Then
            sSQL = "Insert into LIEFEXC select LINR  "
            sSQL = sSQL & ", LIEFBEZ "
            sSQL = sSQL & ", AWERT  "
            sSQL = sSQL & ", ZIELEK "
            sSQL = sSQL & ", FAX "
            sSQL = sSQL & ", KTEXT "
            sSQL = sSQL & ", KUERZEL "
            sSQL = sSQL & ", NOTIZ "
            sSQL = sSQL & ", PLZ "
            sSQL = sSQL & ", STADT "
            sSQL = sSQL & ", STRASSE "
            sSQL = sSQL & ", TEL from LISRT where LINR = " & cLief
            gdBase.Execute sSQL, dbFailOnError
        Else

        End If
        
        reportbildschirm "", "aWKL76a"
        
        loeschNEW "ArtExc", gdBase
    
    End If

Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "LagerPlatzListe"
        Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
        
        Fehlermeldung1
  
End Sub
Private Sub Linienliste()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cLief As String

    If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
    
        loeschNEW "ARTEXC2", gdBase
        CreateTable "ARTEXC2", gdBase
        
        sSQL = "Insert into ARTEXC2 Select KVKPR1,Bestand,MINMEN,LPZ,Artnr,Bezeich,Libesnr,ean  from TOPZ" & srechnertab
        sSQL = sSQL & " order by LPZ,bezeich "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        cLief = ermLief
        loeschNEW "LIEFExc", gdBase
        CreateTable "LIEFEXC", gdBase

        
        
        If cLief <> "0" Then
            sSQL = "Insert into LIEFEXC select LINR  "
            sSQL = sSQL & ", LIEFBEZ "
            sSQL = sSQL & ", AWERT  "
            sSQL = sSQL & ", ZIELEK "
            sSQL = sSQL & ", FAX "
            sSQL = sSQL & ", KTEXT "
            sSQL = sSQL & ", KUERZEL "
            sSQL = sSQL & ", NOTIZ "
            sSQL = sSQL & ", PLZ "
            sSQL = sSQL & ", STADT "
            sSQL = sSQL & ", STRASSE "
            sSQL = sSQL & ", TEL from LISRT where LINR = " & cLief
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        Else

        End If
        
        reportbildschirm "", "aWKL76b"
        
        loeschNEW "ArtExc", gdBase
    
    End If

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Linienliste"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub Preisliste()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cLief As String

    If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
    
        loeschNEW "ARTEXC1", gdBase
        CreateTable "ARTEXC1", gdBase
        
        sSQL = "Insert into ARTEXC1 Select ean,RKZ,gefuehrt,VKPR,KVKPR1,Bestand,LPZ,Artnr,Bezeich,Libesnr  from TOPZ" & srechnertab
        
        sSQL = sSQL & " order by "
        
        If Option1(0).Value = True Then
            sSQL = sSQL & " LPZ,bezeich "
        ElseIf Option1(1).Value = True Then
            sSQL = sSQL & " bezeich "
        End If
        
        
        gdBase.Execute sSQL, dbFailOnError
        
        cLief = ermLief
        loeschNEW "LIEFExc", gdBase
        CreateTable "LIEFEXC", gdBase

        If cLief <> "0" Then
            sSQL = "Insert into LIEFEXC select LINR  "
            sSQL = sSQL & ", LIEFBEZ "
            sSQL = sSQL & ", AWERT  "
            sSQL = sSQL & ", ZIELEK "
            sSQL = sSQL & ", FAX "
            sSQL = sSQL & ", KTEXT "
            sSQL = sSQL & ", KUERZEL "
            sSQL = sSQL & ", NOTIZ "
            sSQL = sSQL & ", PLZ "
            sSQL = sSQL & ", STADT "
            sSQL = sSQL & ", STRASSE "
            sSQL = sSQL & ", TEL from LISRT where LINR = " & cLief
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        reportbildschirm "", "aWKL76c"
        
        loeschNEW "ArtExc", gdBase
    
    End If

Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Preisliste"
        Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
        
        Fehlermeldung1
  
End Sub
Private Sub Konditionenliste()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cLief As String

    If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
    
        loeschNEW "ArtExb", gdBase
        
        sSQL = "Select Konditionen.KONDI,Konditionen.Faktor "
        sSQL = sSQL & ", TOPZ" & srechnertab & ".Bestand"
        sSQL = sSQL & ", TOPZ" & srechnertab & ".LPZ"
        sSQL = sSQL & ", TOPZ" & srechnertab & ".Artnr"
        sSQL = sSQL & ", TOPZ" & srechnertab & ".Bezeich"
        sSQL = sSQL & ", TOPZ" & srechnertab & ".Libesnr "
        sSQL = sSQL & " into ArtExb from TOPZ" & srechnertab
        sSQL = sSQL & " ,Konditionen where konditionen.artnr = TOPZ" & srechnertab & ".artnr"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Delete from ArtExb where Kondi = 0 "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update ArtExb  set ArtExb.Faktor = 1 "
        sSQL = sSQL & " where ArtExb.Faktor = 0 "
        gdBase.Execute sSQL, dbFailOnError
        
        cLief = ermLief
        loeschNEW "LIEFExc", gdBase
        CreateTable "LIEFEXC", gdBase

        If cLief <> "0" Then
            sSQL = "Insert into LIEFEXC select LINR  "
            sSQL = sSQL & ", LIEFBEZ "
            sSQL = sSQL & ", AWERT  "
            sSQL = sSQL & ", ZIELEK "
            sSQL = sSQL & ", FAX "
            sSQL = sSQL & ", KTEXT "
            sSQL = sSQL & ", KUERZEL "
            sSQL = sSQL & ", NOTIZ "
            sSQL = sSQL & ", PLZ "
            sSQL = sSQL & ", STADT "
            sSQL = sSQL & ", STRASSE "
            sSQL = sSQL & ", TEL from LISRT where LINR = " & cLief
            gdBase.Execute sSQL, dbFailOnError
        Else

        End If
        
        reportbildschirm "", "aWKL76d"
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Konditionenliste"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Gewichteliste()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If NewTableSuchenDBKombi("KILOART", gdBase) Then

        If NewTableSuchenDBKombi("TOPZ" & srechnertab, gdBase) Then
            loeschNEW "ArtExb", gdBase
    
            sSQL = "Select GewichtKG "
            sSQL = sSQL & ", TOPZ" & srechnertab & ".AGN "
            sSQL = sSQL & ", TOPZ" & srechnertab & ".Artnr "
            sSQL = sSQL & ", TOPZ" & srechnertab & ".Bezeich"
            sSQL = sSQL & ", TOPZ" & srechnertab & ".Libesnr "
            sSQL = sSQL & ", TOPZ" & srechnertab & ".KVKPR1 "
            sSQL = sSQL & " into ArtExb from TOPZ" & srechnertab
            sSQL = sSQL & " ,KILOART where KILOART.artnr = TOPZ" & srechnertab & ".artnr "
            gdBase.Execute sSQL, dbFailOnError
    
            reportbildschirm "", "aWKL76e"
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Gewichteliste"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermLief()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset

    ermLief = "0"

    sSQL = "Select distinct(LINR) from TOPZ" & srechnertab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If rsrs.RecordCount > 1 Then
        
        Else
            If Not IsNull(rsrs!linr) Then
                ermLief = rsrs!linr
            
            End If
        End If
    
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermLief"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    loeschNEW "ArtExb", gdBase
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


