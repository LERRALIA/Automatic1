VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK25j 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Produktgruppenliste"
   ClientHeight    =   8625
   ClientLeft      =   1755
   ClientTop       =   1845
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
   Icon            =   "frmWK25j.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Auswahlkriterien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12015
      Begin sevCommand3.Command Command2 
         Height          =   405
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Left            =   4680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin sevCommand3.Command Command1 
         Height          =   405
         Left            =   6000
         TabIndex        =   9
         Top             =   720
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   2
         Left            =   9000
         TabIndex        =   11
         Top             =   720
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Schließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Left            =   9000
         TabIndex        =   10
         Top             =   120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Suche Daten"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   20
         Left            =   2520
         TabIndex        =   15
         ToolTipText     =   "Kalender"
         Top             =   240
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   21
         Left            =   2520
         TabIndex        =   16
         ToolTipText     =   "Kalender"
         Top             =   720
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Produktgruppe:"
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
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Lief.Nr.:"
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
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie ein Suchkriterium ein!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   7920
      Width           =   9135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Produktgruppenliste"
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
      TabIndex        =   12
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmWK25j"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            MaskEdBox1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 21        ' Kalender
            MaskEdBox1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
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
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1_KeyUp vbKeyF2, 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR
    
    MaskEdBox1_KeyUp 2, vbKeyF2, 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
'    positionierenwkl25h
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeDialogWK25h() As Long
    On Error GoTo LOKAL_ERROR

    Dim cDatVon As String
    Dim cDatBis As String
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim cLinr As String
    Dim cAgn As String
    Dim bGefunden As Boolean
        
    fnPruefeEingabeDialogWK25h = 0
            
    bGefunden = False
    cDatVon = MaskEdBox1(0).Text
    cDatBis = MaskEdBox1(1).Text
    cLinr = MaskEdBox1(2).Text
    cAgn = Text1.Text
    
    If cDatVon <> "__.__.____" Then
        bGefunden = True
        If Not IsDate(cDatVon) Then
            fnPruefeEingabeDialogWK25h = 1
            Exit Function
        Else
            lDatVon = DateValue(cDatVon)
        End If
    End If
    
    If cDatBis <> "__.__.____" Then
        bGefunden = True
        If Not IsDate(cDatBis) Then
            fnPruefeEingabeDialogWK25h = 2
            Exit Function
        Else
            lDatBis = DateValue(cDatBis)
        End If
    End If
    
    If bGefunden Then
        If cDatVon = "__.__.____" Or cDatBis = "__.__.____" Then
            fnPruefeEingabeDialogWK25h = 3
            Exit Function
        End If
    End If
    
    If lDatVon > lDatBis Then
        fnPruefeEingabeDialogWK25h = 4
        Exit Function
    End If
    
    If cLinr <> "______" Then
        bGefunden = True
    End If
    
    If cAgn <> "" Then
        bGefunden = True
    End If
    
    If Not bGefunden Then
        fnPruefeEingabeDialogWK25h = 5
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK25h"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub SucheDatenWK25j()
    On Error GoTo LOKAL_ERROR

    Dim cDatVon As String
    Dim cDatBis As String
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim cLinr As String
    Dim cPGN As String
    Dim lagn As Long
    Dim cSQL As String
    Dim cLBSatz As String
    Dim cFeld As String
    Dim dWert As Double
    Dim rsrs As Recordset
    Dim rsagn As Recordset
    Dim bAnd As Boolean
    
    Dim lSumAnzahl As Long
    Dim lAnzPos As Long
    Dim dSumUmsatz As Double
    
    anzeigeNew "normal", "Daten werden ermittelt...", Label6
    Screen.MousePointer = 11
    
    cDatVon = MaskEdBox1(0).Text
    cDatBis = MaskEdBox1(1).Text
    cLinr = MaskEdBox1(2).Text
    
    cPGN = Text1.Text

    If cDatBis <> "__.__.____" Then
        lDatBis = DateValue(cDatBis)
    End If
    
    loeschNEW "PGNte", dabalokal
    CreateTable "PGNTE", dabalokal
    
    loeschNEW "PGNteK", dabalokal
    CreateTable "PGNTEK", dabalokal
    
    cSQL = "INSERT Into PGNTEK (von,bis,linr,PGN)  values "
    cSQL = cSQL & "( '" & Trim(MaskEdBox1(0).Text) & "','" & Trim(MaskEdBox1(1).Text) & "' "
    
    If cLinr = "______" Then
        cSQL = cSQL & ", 'alle' "
    Else
        cSQL = cSQL & ", '" & cLinr & "' "
    End If
    
    If cPGN = "" Then
        cSQL = cSQL & ", 'alle' "
    Else
        cSQL = cSQL & ", '" & cPGN & "' "
    End If
    
    cSQL = cSQL & " ) "
    
    
    
    dabalokal.Execute cSQL, dbFailOnError
    
    
    cSQL = "INSERT Into PGNTE Select  artikel.PGN "
    cSQL = cSQL & " , Kassjour.artnr "
    cSQL = cSQL & ", Kassjour.bezeich "
    cSQL = cSQL & ", Kassjour.menge "
    cSQL = cSQL & ", Kassjour.preis "
    cSQL = cSQL & ", Kassjour.adate "
    cSQL = cSQL & ", Kassjour.azeit "
    cSQL = cSQL & ", Kassjour.linr "
    cSQL = cSQL & ", pgndbf.pgnbezeich,lisrt.liefbez"
    cSQL = cSQL & " from KASSJOUR, pgndbf, lisrt , artikel where kassjour.artnr = artikel.artnr and "
    cSQL = cSQL & " pgndbf.pgn = artikel.pgn and "
    cSQL = cSQL & " kassjour.linr = lisrt.linr and "
    If cDatVon <> "__.__.____" Then
        lDatVon = DateValue(cDatVon)
        cSQL = cSQL & "KASSJOUR.ADATE >= " & Trim$(Str$(lDatVon)) & " "
        bAnd = True
    End If
    
    If cDatBis <> "__.__.____" Then
        lDatBis = DateValue(cDatBis)
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "KASSJOUR.ADATE <= " & Trim$(Str$(lDatBis)) & " "
        bAnd = True
    End If
    
    If cLinr <> "______" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "KASSJOUR.LINR = " & cLinr & " "
        bAnd = True
    End If
    
    If cPGN <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "artikel.PGN = " & cPGN & " "
        bAnd = True
    End If

    cSQL = cSQL & " and Artikel.pgn <> 0 "
    dabalokal.Execute cSQL, dbFailOnError
    

    If Datendrin("PGNTE", dabalokal) Then
        anzeigeNew "normal", "Druckvorschau wird erstellt...", Label6
        reportbildschirmApp "WKL033", "aWKL25Ja"
        
    Else
        anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
    End If
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenWK25j"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MaskEdBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 2     'LiNr
                gF2Prompt.cFeld = "LINR"
        End Select
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                ctmp = gF2Prompt.cWahl
                Select Case Index
                    Case Is = 2
                        ctmp = ctmp & String$(6 - Len(ctmp), "_")
                    Case Is = 3
                        ctmp = ctmp & String$(3 - Len(ctmp), "_")
                End Select
                MaskEdBox1(Index).Text = ctmp
            End If
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lRet As Long
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'Suchen
            lRet = fnPruefeEingabeDialogWK25h()
            Select Case lRet
                Case Is = 0
                    anzeigeNew "normal", "Datenbank wird aktualisiert...", Label6
                    Datenbankwechsel
                    Me.Refresh
                    SucheDatenWK25j
                Case Is = 1
                    anzeigeNew "rot", "Das eingegebene VON-Datum ist ungültig!", Label6
                    
                    MaskEdBox1(0).SetFocus
                Case Is = 2
                    anzeigeNew "rot", "Das eingegebene BIS-Datum ist ungültig!", Label6
                    MaskEdBox1(1).SetFocus
                Case Is = 3
                    anzeigeNew "rot", "Bitte VON- und BIS-Datum eingeben!", Label6
                    
                    MaskEdBox1(0).SetFocus
                Case Is = 4
                    anzeigeNew "rot", "Das VON-Datum ist größer als das BIS-Datum!", Label6
                    
                    MaskEdBox1(0).SetFocus
                Case Is = 5
                    anzeigeNew "rot", "Bitte mindestens ein Auswahlkriterium eingeben!", Label6
                   
                    MaskEdBox1(0).SetFocus
            End Select
            
        
        Case Is = 2     'Schließen
            Unload frmWK25j
    End Select
        
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        gF2Prompt.cFeld = "PGN"
                
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text1.Text = gF2Prompt.cWahl
            End If
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppenliste ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

