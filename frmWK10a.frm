VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK10a 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Terminierte Kassenverkaufspreise"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWK10a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'Kein
      Caption         =   "Termin für Kassen-VK-Preis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7920
      TabIndex        =   6
      Top             =   840
      Width           =   3975
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1680
         Width           =   1335
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   2
         Left            =   2640
         TabIndex        =   21
         Top             =   3120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Schließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Speichern"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   18
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   4
         Left            =   2640
         TabIndex        =   27
         ToolTipText     =   "Kalender"
         Top             =   2160
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   5
         Left            =   2640
         TabIndex        =   28
         ToolTipText     =   "Kalender"
         Top             =   2640
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Termin für Kassen-VK-Preis"
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
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Datum bis:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Datum von:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "neuer Preis:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "alter Preis:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Bez.:"
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
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "ArtNr:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'Kein
      Caption         =   "Artikel mit terminierten Preisen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   11895
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   10455
      End
      Begin VB.ListBox List3 
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
         TabIndex        =   4
         Top             =   360
         Width           =   10455
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   3
         Left            =   10680
         TabIndex        =   22
         Top             =   360
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Löschen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Artikel mit terminierten Preisen"
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
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Caption         =   "Artikel laut Vorauswahl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7935
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
         Height          =   3000
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Kalendar"
         Top             =   600
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
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label3 
         Caption         =   "Artikel laut Vorauswahl"
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
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminierte Kassenverkaufspreise"
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
      TabIndex        =   23
      Top             =   120
      Width           =   10935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWK10a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim iRet As Long
    
    Select Case Index
        
        Case Is = 4     'kalender
            
            MaskEdBox1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            
        Case Is = 5     'kalender
            MaskEdBox1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 0
    
    WK10aPositionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    List1.Clear
    List1.AddItem "ArtNr. Artikelbezeichnung                   Kassen-VK"
    
    List3.Clear
    List3.AddItem "ArtNr. Artikelbezeichnung                    alter VK   neuer VK Datum Von  Datum Bis  Aktiv"
    
    Label2(0).Caption = ""
    Label2(1).Caption = ""
    Label2(2).Caption = ""
    
    Text1.Text = ""
    MaskEdBox1(0).Text = "__.__.____"
    MaskEdBox1(1).Text = "__.__.____"
    
    HoleArtikelVorauswahlWK10a
    
    HoleArtikelTerminPreiseWK10a
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1

End Sub
Private Sub WK10aPositionieren()
On Error GoTo LOKAL_ERROR
    Frame1.Top = 840
    Frame1.Left = 0
    Frame1.Height = 3735
    Frame1.Width = 7935
    Frame1.Visible = True
    
    Frame2.Top = 4440
    Frame2.Left = 0
    Frame2.Height = 4215
    Frame2.Width = 11895
    Frame2.Visible = True
    
    Frame3.Top = 840
    Frame3.Left = 7920
    Frame3.Height = 3735
    Frame3.Width = 3975
    Frame3.Visible = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WK10aPositionieren"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeDialogWK10a() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim dWert As Double
    Dim lVon As Long
    Dim lBis As Long
    Dim lVonDatei As Long
    Dim lBisDatei As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnPruefeEingabeDialogWK10a = 0
    
    'Artikel ausgewählt?
    cFeld = Label2(0).Caption
    If cFeld = "" Then
        fnPruefeEingabeDialogWK10a = 1
        Exit Function
    End If
    
    'Preis korrekt?
    cFeld = Text1.Text
    If cFeld = "" Then
        fnPruefeEingabeDialogWK10a = 2
        Exit Function
    End If
    
    cFeld = fnMoveComma2Point$(cFeld)
    dWert = Val(cFeld)
    If dWert >= 100000 Then
        fnPruefeEingabeDialogWK10a = 3
        Exit Function
    End If
            
    cFeld = Format$(dWert, "######0.00")
    Text1.Text = cFeld
    
    
    'VON-Datum eingegeben?
    cFeld = MaskEdBox1(0).Text
    If Not IsDate(cFeld) Then
        fnPruefeEingabeDialogWK10a = 4
        Exit Function
    End If
    lVon = DateValue(cFeld)
    
    'BIS-Datum eingegeben?
    cFeld = MaskEdBox1(1).Text
    If Not IsDate(cFeld) Then
        fnPruefeEingabeDialogWK10a = 5
        Exit Function
    End If
    lBis = DateValue(cFeld)
    
    If lVon > lBis Then
        fnPruefeEingabeDialogWK10a = 6
        Exit Function
    End If
    
    
    'auf überschneidenden Zeitraum prüfen
    cSQL = "Select * from PRSTERM where "
    cSQL = cSQL & "ARTNR = " & Trim$(Label2(0).Caption) & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!DAT_VON) Then
                lVonDatei = rsrs!DAT_VON
            Else
                lVonDatei = -1
            End If
            If Not IsNull(rsrs!DAT_BIS) Then
                lBisDatei = rsrs!DAT_BIS
            Else
                lBisDatei = -1
            End If
            
            If lVonDatei <> -1 And lBisDatei <> -1 Then
                                    
                'VON liegt in einem bestehenden Zeitraum
                If lVon >= lVonDatei And lVon < lBisDatei Then
                    fnPruefeEingabeDialogWK10a = 7
                    Exit Function
                End If
                
                'BIS liegt in einem bestehenden Zeitraum
                If lBis > lVonDatei And lBis <= lBisDatei Then
                    fnPruefeEingabeDialogWK10a = 7
                    Exit Function
                End If
                
                'VON und BIS umrahmen bestehenden Zeitraum
                If lVon <= lVonDatei And lBis >= lBisDatei Then
                    fnPruefeEingabeDialogWK10a = 7
                    Exit Function
                End If
                
                'VON und BIS werden von bestehendem Zeitraum umrahmt
                If lVon >= lVonDatei And lBis <= lBisDatei Then
                    fnPruefeEingabeDialogWK10a = 7
                    Exit Function
                End If
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Function
Private Sub HoleArtikelTerminPreiseWK10a()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim cLBSatz As String
    Dim cFeld As String
    Dim lartnr As Long
    Dim dWert As Double
    
    List4.Clear
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    rsArt.Index = "ARTNR"
    
    cSQL = "Select * from PRSTERM order by DAT_VON, ARTNR "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLBSatz = ""
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            lartnr = Val(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
                
            rsArt.Seek "=", lartnr
            If Not rsArt.NoMatch Then
                If Not IsNull(rsArt!BEZEICH) Then
                    cFeld = rsArt!BEZEICH
                Else
                    cFeld = ""
                End If
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!KVKPR1ALT) Then
                dWert = rsrs!KVKPR1ALT
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "######0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!KVKPR1NEU) Then
                dWert = rsrs!KVKPR1NEU
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "######0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!DAT_VON) Then
                dWert = rsrs!DAT_VON
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "DD.MM.YYYY")
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!DAT_BIS) Then
                dWert = rsrs!DAT_BIS
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "DD.MM.YYYY")
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!Status) Then
                dWert = rsrs!Status
            Else
                dWert = 0
            End If
            If dWert = 0 Then
                cFeld = "N"
            Else
                cFeld = "J"
            End If
            cLBSatz = cLBSatz & cFeld & " "
            
            List4.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsArt.Close: Set rsArt = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleArtikelTerminPreiseWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub
Private Sub HoleArtikelVorauswahlWK10a()
    On Error GoTo LOKAL_ERROR
            
    Dim rsrs        As Recordset
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim cKVkPr1     As String
    
    List2.Clear
    
    If NewTableSuchenDBKombi("TOP" & srechnertab, gdBase) = False Then
        Exit Sub
    End If
    
    Set rsrs = gdBase.OpenRecordset("TOP" & srechnertab)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
                cArtNr = Space$(6 - Len(cArtNr)) & cArtNr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                    cBezeich = cBezeich & Space$(35 - Len(cBezeich))
                End If
                
                If Not IsNull(rsrs!KVKPR1) Then
                    cKVkPr1 = rsrs!KVKPR1
                    cKVkPr1 = Space$(10 - Len(Format$(cKVkPr1, "######0.00"))) & Format$(cKVkPr1, "######0.00")
                End If
                
                cLBSatz = cArtNr & " " & cBezeich & " " & cKVkPr1
                List2.AddItem cLBSatz
                
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleArtikelVorauswahlWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub InsertTerminPreisWK10a()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cFeld As String
    
    Dim lartnr As Long
    Dim dKVkPr1Alt As Double
    Dim dKVkPr1Neu As Double
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    cFeld = Label2(0).Caption
    lartnr = Val(cFeld)
    
    cFeld = Label2(2).Caption
    cFeld = fnMoveComma2Point$(cFeld)
    dKVkPr1Alt = Val(cFeld)
    
    cFeld = Text1.Text
    cFeld = fnMoveComma2Point$(cFeld)
    dKVkPr1Neu = Val(cFeld)
    
    cFeld = MaskEdBox1(0).Text
    lDatVon = DateValue(cFeld)
    
    cFeld = MaskEdBox1(1).Text
    lDatBis = DateValue(cFeld)
    
    cSQL = "Select * from PRSTERM where ARTNR = -1"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    rsrs.AddNew
    rsrs!artnr = lartnr
    rsrs!KVKPR1ALT = dKVkPr1Alt
    rsrs!KVKPR1NEU = dKVkPr1Neu
    rsrs!DAT_VON = lDatVon
    rsrs!DAT_BIS = lDatBis
    rsrs!Status = 0
    rsrs.Update
    
    rsrs.Close: Set rsrs = Nothing
    
    
    HoleArtikelTerminPreiseWK10a
    
    SSCommand1_Click 1
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertTerminPreisWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub
Private Sub LoescheTerminPreisWK10a()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cArtNr As String
    Dim cFeld As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cSQL As String
    
    cLBSatz = List4.list(List4.ListIndex)
    cArtNr = Mid(cLBSatz, 1, 6)
    cFeld = Mid(cLBSatz, 66, 10)
    lVon = DateValue(cFeld)
    cFeld = Mid(cLBSatz, 77, 10)
    lBis = DateValue(cFeld)
    
    cSQL = "Delete from PRSTERM where "
    cSQL = cSQL & "ARTNR = " & cArtNr & " "
    cSQL = cSQL & "and DAT_VON = " & Trim$(Str$(lVon)) & " "
    cSQL = cSQL & "and DAT_BIS = " & Trim$(Str$(lBis)) & " "
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    HoleArtikelTerminPreiseWK10a
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheTerminPreisWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub
Private Sub SchreibeTerminPreisWK10a()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet As Long
    
    '*** Eingabeprüfung ***
    lRet = fnPruefeEingabeDialogWK10a()
    Select Case lRet
        Case Is = 0
            InsertTerminPreisWK10a
        Case Is = 1
            MsgBox "Kein Artikel gewählt!", vbCritical, "STOP!"
        Case Is = 2
            MsgBox "Kein Preis eingegeben!", vbCritical, "STOP!"
            Text1.SetFocus
        Case Is = 3
            MsgBox "Zu hohen Preis eingegeben!", vbCritical, "STOP!"
            Text1.SetFocus
        Case Is = 4
            MsgBox "Fehlendes oder ungültiges VON-Datum!", vbCritical, "STOP!"
            MaskEdBox1(0).SetFocus
        Case Is = 5
            MsgBox "Fehlendes oder ungültiges BIS-Datum!", vbCritical, "STOP!"
            MaskEdBox1(1).SetFocus
        Case Is = 6
            MsgBox "Das VON-Datum ist größer als das BIS-Datum!", vbCritical, "STOP!"
            MaskEdBox1(0).SetFocus
        Case Is = 7
            MsgBox "Zeitraumüberschneidung mit bereits vorhandenem Terminpreis!", vbCritical, "STOP!"
            MaskEdBox1(0).SetFocus
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTerminPreisWK10a"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub



Private Sub List2_Click()
    On Error GoTo LOKAL_ERROR

    Dim cLBSatz As String
    
    cLBSatz = List2.list(List2.ListIndex)
    
    Label2(0).Caption = Mid(cLBSatz, 1, 6)
    Label2(1).Caption = Mid(cLBSatz, 8, 35)
    Label2(2).Caption = Trim$(Mid(cLBSatz, 45, 10))

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub MaskEdBox1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = glSelBack1
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1

End Sub
Private Sub MaskEdBox1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Long
    
    Select Case Index
        Case Is = 0     'Speichern
            SchreibeTerminPreisWK10a
            
        Case Is = 1     'Leeren
            Label2(0).Caption = ""
            Label2(1).Caption = ""
            Label2(2).Caption = ""
            Text1.Text = ""
            MaskEdBox1(0).Text = "__.__.____"
            MaskEdBox1(1).Text = "__.__.____"
            
        Case Is = 2     'Schließen
            Unload frmWK10a
            
        Case Is = 3     'Löschen
            If List4.ListIndex < 0 Then
                MsgBox "Bitte einen Eintrag auswählen!", vbCritical, "STOP!"
                List4.SetFocus
            Else
                iRet = MsgBox("Wollen Sie den ausgewählten Datensatz wirklich löschen?", vbYesNo + vbQuestion, "LÖSCHEN")
                If iRet = vbYes Then
                    LoescheTerminPreisWK10a
                End If
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Terminierte Kassenverkaufspreise auf. "
    
    Fehlermeldung1
End Sub


