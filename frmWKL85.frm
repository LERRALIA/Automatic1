VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL85 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Spezialetikettendruck"
   ClientHeight    =   8940
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   12210
   Icon            =   "frmWKL85.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8940
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximiert
   Begin sevCommand3.Command cmdRed 
      Height          =   495
      Left            =   10200
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "rote Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "schwarze Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "für Etikettendruck vorgesehene Artikel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin MSComDlg.CommonDialog cdlprinter 
         Left            =   5160
         Top             =   7080
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bezeich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   12
         Top             =   7800
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LiefNr, Bezeich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   11
         Top             =   7440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LiefNr, LiefBestNr, Bezeich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   8160
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LiefNr, Linie, Bezeich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   7800
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ArtNr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   7440
         Width           =   2295
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
         Height          =   6150
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   2
         Top             =   720
         Width           =   7695
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7695
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7080
         Top             =   7320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Left            =   4920
         TabIndex        =   13
         Top             =   8040
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Daten neu sortieren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Neu sortieren nach:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl Etiketten:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl Artikel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Index           =   5
      Left            =   8040
      TabIndex        =   18
      Top             =   8040
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Schließen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Index           =   4
      Left            =   8040
      TabIndex        =   17
      Top             =   7320
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Drucke Etiketten"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Index           =   2
      Left            =   8040
      TabIndex        =   16
      Top             =   1320
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Lösche alle Sätze"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   15
      Top             =   720
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Lösche markierte Sätze"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   14
      Top             =   120
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Markiere alle Sätze"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmWKL85"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bRot         As Boolean
Public bSchwarz     As Boolean
Private Function fnSucheMarkierteEintraegeWED23() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnSucheMarkierteEintraegeWED23 = 0
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            fnSucheMarkierteEintraegeWED23 = 1
            Exit For
        End If
    Next lcount
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnSucheMarkierteEintraegeWED23"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub FuelleEtikettenListeRotWED23(lSort As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cOrderBy As String
    Dim lWert As Long
    Dim dWert As Double
    Dim cFeld As String
    Dim cBezeich As String
    Dim cLBSatz As String
    Dim lRet As Long
    Dim lAnzArt As Long
    Dim lAnzEti     As Long
    Dim lartnr      As Long
    Dim bWeiter     As Boolean
    Dim ctmp        As String
    Dim sSQL        As String
    
    
    loeschNEW "temp", gdBase
  
    sSQL = " Select etidru.artnr, etidru.bezeich, etidru.ean, artikel.vkpr as VKPR,etidru.vkpr as KVKPR1  into Temp from ETIDRU,artikel  where ANZAHL > 0 and etidru.VKPR < artikel.VKpr and etidru.artnr=artikel.artnr "
    cSQL = "Select etidru.*,artikel.vkpr as vkreal from ETIDRU,artikel where etidru.ANZAHL > 0 and etidru.VKPR < artikel.vkpr and artikel.artnr=etidru.artnr "
    
    Select Case lSort
        Case Is = 0
            cOrderBy = "order by etidru.FILNR, etidru.ARTNR"
        Case Is = 1
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.LPZ, etidru.BEZEICH"
        Case Is = 2
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.LIBESNR, etidru.BEZEICH"
        Case Is = 3
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.BEZEICH"
        Case Is = 4
            cOrderBy = "order by etidru.FILNR, etidru.BEZEICH"
    End Select
    
    sSQL = sSQL & cOrderBy
    
    cSQL = cSQL & cOrderBy
    
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    List2.Clear
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzArt = 0
        lAnzEti = 0
        Do While Not rsrs.EOF
            lAnzArt = lAnzArt + 1
            cLBSatz = ""
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = -1
            End If
            lartnr = lWert
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!filnr) Then
                lWert = rsrs!filnr
            Else
                lWert = -1
            End If
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cBezeich = cFeld
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = -1
            End If
            If dWert > 0 Then
                cFeld = Format$(dWert, "######0.00")
            Else
                cFeld = "0,00"
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!ANZAHL) Then
                lWert = rsrs!ANZAHL
            Else
                lWert = 0
            End If
            
            bWeiter = True
            If lWert >= 1000 Then
                ctmp = "Von Artikel " & Trim$(Str$(lartnr)) & " / " & cBezeich & " " & vbCrLf
                ctmp = ctmp & "sollen " & Trim$(Str$(lWert)) & " Etiketten gedruckt werden." & vbCrLf & vbCrLf
                ctmp = ctmp & "Ist das wirklich gewünscht?"
                lRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage")
                If lRet = vbYes Then
                    bWeiter = True
                Else
                    bWeiter = False
                End If
            End If
                
            If bWeiter Then
                lAnzEti = lAnzEti + lWert
                If dWert > 0 Then
                    cFeld = Trim$(Str$(lWert))
                Else
                    cFeld = "0"
                End If
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                    
                List2.AddItem cLBSatz
            End If
            rsrs.MoveNext
        Loop
        rsrs.Close: Set rsrs = Nothing
        
        Label2(0).Caption = Trim$(Str$(lAnzArt))
        Label2(0).Refresh
        Label2(1).Caption = Trim$(Str$(lAnzEti))
        Label2(1).Refresh
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleEtikettenListeRotWED23"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleEtikettenListeBlackWED23(lSort As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    Dim cOrderBy    As String
    Dim lWert       As Long
    Dim dWert       As Double
    Dim cFeld       As String
    Dim cBezeich    As String
    Dim cLBSatz     As String
    Dim lRet        As Long
    Dim lAnzArt     As Long
    Dim lAnzEti     As Long
    Dim lartnr      As Long
    Dim bWeiter     As Boolean
    Dim ctmp        As String
    Dim sSQL        As String
    
    loeschNEW "TEMP", gdBase
    
    sSQL = " Select etidru.artnr, etidru.bezeich, etidru.ean, artikel.vkpr as VKPR,etidru.vkpr as KVKPR1  into Temp from ETIDRU,artikel  where ANZAHL > 0 and etidru.VKPR = artikel.VKpr and etidru.artnr=artikel.artnr "
    cSQL = "Select etidru.*,artikel.vkpr as vkreal from ETIDRU,artikel where etidru.ANZAHL > 0 and etidru.VKPR = artikel.vkpr and artikel.artnr=etidru.artnr "
    
    Select Case lSort
        Case Is = 0
            cOrderBy = "order by etidru.FILNR, etidru.ARTNR"
        Case Is = 1
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.LPZ, etidru.BEZEICH"
        Case Is = 2
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.LIBESNR, etidru.BEZEICH"
        Case Is = 3
            cOrderBy = "order by etidru.FILNR, etidru.LINR, etidru.BEZEICH"
        Case Is = 4
            cOrderBy = "order by etidru.FILNR, etidru.BEZEICH"
    End Select
    
    cSQL = cSQL & cOrderBy
    sSQL = sSQL & cOrderBy
    
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    List2.Clear
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzArt = 0
        lAnzEti = 0
        Do While Not rsrs.EOF
            lAnzArt = lAnzArt + 1
            cLBSatz = ""
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = -1
            End If
            lartnr = lWert
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!filnr) Then
                lWert = rsrs!filnr
            Else
                lWert = -1
            End If
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cBezeich = cFeld
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = -1
            End If
            If dWert > 0 Then
                cFeld = Format$(dWert, "######0.00")
            Else
                cFeld = "0,00"
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!ANZAHL) Then
                lWert = rsrs!ANZAHL
            Else
                lWert = 0
            End If
            
            bWeiter = True
            If lWert >= 1000 Then
                ctmp = "Von Artikel " & Trim$(Str$(lartnr)) & " / " & cBezeich & " " & vbCrLf
                ctmp = ctmp & "sollen " & Trim$(Str$(lWert)) & " Etiketten gedruckt werden." & vbCrLf & vbCrLf
                ctmp = ctmp & "Ist das wirklich gewünscht?"
                lRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
                If lRet = vbYes Then
                    bWeiter = True
                Else
                    bWeiter = False
                End If
            End If
                
            If bWeiter Then
                lAnzEti = lAnzEti + lWert
                If dWert > 0 Then
                    cFeld = Trim$(Str$(lWert))
                Else
                    cFeld = "0"
                End If
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                    
                List2.AddItem cLBSatz
            End If
            rsrs.MoveNext
        Loop
        rsrs.Close: Set rsrs = Nothing
        Label2(0).Caption = Trim$(Str$(lAnzArt))
        Label2(0).Refresh
        Label2(1).Caption = Trim$(Str$(lAnzEti))
        Label2(1).Refresh
        
    End If
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleEtikettenListeBlackWED23"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleEtikettenListeWED23(lSort As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cOrderBy As String
    Dim lWert As Long
    Dim dWert As Double
    Dim cFeld As String
    Dim cBezeich As String
    Dim cLBSatz As String
    Dim lRet As Long
    Dim lAnzArt As Long
    Dim lAnzEti As Long
    Dim lartnr As Long
    Dim bWeiter As Boolean
    Dim ctmp As String
    
    cSQL = "Select * from ETIDRU where ANZAHL > 0 "
    
    Select Case lSort
        Case Is = 0
            cOrderBy = "order by FILNR, ARTNR"
            
        Case Is = 1
            cOrderBy = "order by FILNR, LINR, LPZ, BEZEICH"
            
        Case Is = 2
            cOrderBy = "order by FILNR, LINR, LIBESNR, BEZEICH"
            
        Case Is = 3
            cOrderBy = "order by FILNR, LINR, BEZEICH"
            
        Case Is = 4
            cOrderBy = "order by FILNR, BEZEICH"
            
    End Select
    
    cSQL = cSQL & cOrderBy
    
    List2.Clear
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzArt = 0
        lAnzEti = 0
        Do While Not rsrs.EOF
            lAnzArt = lAnzArt + 1
            cLBSatz = ""
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = -1
            End If
            lartnr = lWert
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!filnr) Then
                lWert = rsrs!filnr
            Else
                lWert = -1
            End If
            If lWert > 0 Then
                cFeld = Trim$(Str$(lWert))
            Else
                cFeld = ""
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cBezeich = cFeld
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = -1
            End If
            If dWert > 0 Then
                cFeld = Format$(dWert, "######0.00")
            Else
                cFeld = "0,00"
            End If
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            If Not IsNull(rsrs!ANZAHL) Then
                lWert = rsrs!ANZAHL
            Else
                lWert = 0
            End If
            
            bWeiter = True
            If lWert >= 1000 Then
                ctmp = "Von Artikel " & Trim$(Str$(lartnr)) & " / " & cBezeich & " " & vbCrLf
                ctmp = ctmp & "sollen " & Trim$(Str$(lWert)) & " Etiketten gedruckt werden." & vbCrLf & vbCrLf
                ctmp = ctmp & "Ist das wirklich gewünscht?"
                lRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
                If lRet = vbYes Then
                    bWeiter = True
                Else
                    bWeiter = False
                End If
            End If
                
            If bWeiter Then
                lAnzEti = lAnzEti + lWert
                If dWert > 0 Then
                    cFeld = Trim$(Str$(lWert))
                Else
                    cFeld = "0"
                End If
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                    
                List2.AddItem cLBSatz
            End If
            rsrs.MoveNext
        Loop
        rsrs.Close: Set rsrs = Nothing
        
        Label2(0).Caption = Trim$(Str$(lAnzArt))
        Label2(0).Refresh
        Label2(1).Caption = Trim$(Str$(lAnzEti))
        Label2(1).Refresh
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleEtikettenListeWED23"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LoescheSatzWED23(cLBSatz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lartnr As Long
    Dim lFilNr As Long
    
    If cLBSatz = "ALLE" Then
        cSQL = "Delete from ETIDRU"
    Else
        lartnr = Val(Trim$(Left$(cLBSatz, 6)))
        lFilNr = Val(Trim$(Mid$(cLBSatz, 8, 3)))
        
        cSQL = "Delete from ETIDRU "
        cSQL = cSQL & "where ARTNR = " & Trim$(Str$(lartnr)) & " "
        cSQL = cSQL & "and FILNR = " & Trim$(Str$(lFilNr)) & " "
    End If
    
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheSatzWED23"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cmdRed_Click()
On Error GoTo LOKAL_ERROR

    Dim lSort As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    For lcount = 0 To 4
        If Option1(lcount).Value = True Then
            lSort = lcount
            Exit For
        End If
    Next lcount
        
    FuelleEtikettenListeRotWED23 lSort
    bRot = True
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdRed_Click"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    Dim lSort As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    For lcount = 0 To 4
        If Option1(lcount).Value = True Then
            lSort = lcount
            Exit For
        End If
    Next lcount
    
    FuelleEtikettenListeBlackWED23 lSort
    bSchwarz = True
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    List1.Clear
    List2.Clear
    List1.AddItem "ArtNr. Fil Artikelbezeichnung                    VK-Preis    Anz.Eti"
    
    FuelleEtikettenListeWED23 1
    bSchwarz = False
    bRot = False
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
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

Private Sub SSCommand1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lSort As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    For lcount = 0 To 4
        If Option1(lcount).Value = True Then
            lSort = lcount
            Exit For
        End If
    Next lcount
        
    FuelleEtikettenListeWED23 lSort
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub SSCommand2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim lcount          As Long
    Dim lRet            As Long
    Dim ctmp            As String
    Dim cPfad           As String
    Dim cSQL            As String
    
    cPfad = gcPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'Markiere alle Sätze
            For lcount = List2.ListCount - 1 To 0 Step -1
                List2.Selected(lcount) = True
            Next lcount
            
        Case Is = 1     'Lösche markierte Sätze
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    cLBSatz = List2.list(lcount)
                    LoescheSatzWED23 cLBSatz
                End If
            Next lcount
            SSCommand1_Click
            
        Case Is = 2     'Lösche alle Sätze
            LoescheSatzWED23 "ALLE"
            SSCommand1_Click
            
        Case Is = 4     'Drucke Etiketten
        
            Kill cPfad & "Temp.dbf"
            cSQL = "Select * into Temp IN '" & cPfad & "' 'dbase IV;' from Temp "
            gdBase.Execute cSQL, dbFailOnError
          
            If bRot Then
                cdlprinter.ShowPrinter
                
                reportbildschirmohneDrucker "", "naegr"
                bRot = False
                
            ElseIf bSchwarz Then
                cdlprinter.ShowPrinter
                reportbildschirmohneDrucker "", "naegs"
                bSchwarz = False
                
            End If

        Case Is = 5     'Schließen
            Unload frmWKL85
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SSCommand2_Click"
        Fehler.gsFehlertext = "Es ist im Programmteil Spezialetiketten ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub


