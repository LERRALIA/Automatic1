VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL136 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "allgemeine Kassenvorgänge"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Caption         =   "Was möchten Sie sich anschauen?"
      Height          =   4935
      Left            =   8160
      TabIndex        =   15
      Top             =   1080
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "alte Gutscheine"
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   29
         Top             =   4200
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Coupondaten"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   28
         Top             =   3840
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Kreditkartenzahlungen"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   3480
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Preisänderung an der Kasse"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   3120
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "PLZ Einzugsgebiet"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   2760
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Kundenauslieferung"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Verkäufe mit AboPlus-Karte"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Kredittilgung"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bediener der Kasse"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2. Bediener bei Stornos"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ein- und Auszahlungen"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Protokollart"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "Vormonat"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "aktueller Monat"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gestern"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Heute"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Datum Voreinstellung"
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
         TabIndex        =   14
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   4
      Top             =   360
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
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
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   0
      Left            =   1680
      TabIndex        =   24
      ToolTipText     =   "Kalender"
      Top             =   960
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
      Index           =   1
      Left            =   3720
      TabIndex        =   25
      ToolTipText     =   "Kalender"
      Top             =   960
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
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "bis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "von:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "allgemeine Kassenvorgänge"
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
      TabIndex        =   2
      Top             =   120
      Width           =   9135
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
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL136"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            Text1(1).SetFocus
        Case Is = 1
            Text1(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub

Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
     
    Select Case Index
        Case 11
            gsHelpstring = "allgemeine Kassenvorgänge"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL136
        Case 1 'suchen
            If Option2(0).Value = True Then
                SucheDaten
            ElseIf Option2(1).Value = True Then
                SucheDaten2BedBeiStorno
            ElseIf Option2(2).Value = True Then
                SucheDatenBedKasse
            ElseIf Option2(3).Value = True Then
                SucheDatenKedittilgung
            ElseIf Option2(4).Value = True Then
                
                SucheDatenAboPlus
                ExportiereAboPlus
                
            ElseIf Option2(5).Value = True Then
                SucheDatenKunden_Auslieferung
            ElseIf Option2(6).Value = True Then
                SucheDatenPLZ_Erfassung
            ElseIf Option2(7).Value = True Then
                SucheDatenPreisÄnderung_Kasse
            ElseIf Option2(8).Value = True Then
                SucheDatenKeditkartenzahlung
            ElseIf Option2(9).Value = True Then
                SucheDatenCouponEinloesungen
            ElseIf Option2(10).Value = True Then
                SucheDatenAlteGutscheine
            Else
                anzeige "rot", "Wählen Sie die Protokollart aus!", Label1(4)
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDatenPreisÄnderung_Kasse()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "PREISEDITKASSE_PRINT", gdBase
    CreateTableT2 "PREISEDITKASSE_PRINT", gdBase
    
    sSQL = "Insert into PREISEDITKASSE_PRINT Select * from PREISEDITKASSE"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("PREISEDITKASSE_PRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

        reportbildschirm "", "aWKL136e"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenPreisÄnderung_Kasse"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenPLZ_Erfassung()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "PLZPRINT", gdBase
    CreateTableT2 "PLZPRINT", gdBase
    
    sSQL = "Insert into PLZPRINT Select * from PLZGEBIET"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PLZPRINT Set PLZ = 'Ausland' "
    sSQL = sSQL & " where PLZ = '99999'"
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("PLZPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

        reportbildschirm "", "aWKL136d"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenPLZ_Erfassung"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenKeditkartenzahlung()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "KKZAHLTE_PRINT", gdBase
    CreateTableT2 "KKZAHLTE_PRINT", gdBase

    sSQL = "Insert into KKZAHLTE_PRINT Select * from KKZAHL "
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("KKZAHLTE_PRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKLKK"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenKeditkartenzahlung"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenAlteGutscheine()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "ALTERG_PRINT", gdBase
    CreateTableT2 "ALTERG_PRINT", gdBase

    sSQL = "Insert into ALTERG_PRINT Select * from ALTERG "
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("ALTERG_PRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKLAG"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenAlteGutscheine"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenCouponEinloesungen()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    Dim cBudniKundNr As String
    Dim lCouponLinr As Long
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    

    Dim rsLi            As DAO.Recordset
    
    cBudniKundNr = ""
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDIBUDNI' "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        cBudniKundNr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(cBudniKundNr) = 0 Then
    
        anzeige "rot", "Es kann keine Budni-Kundennummer ermittelt werden.", Label1(4)
        Exit Sub



    End If
    
    lCouponLinr = 0
    sSQL = "Select linr from COUPE "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        rsLi.MoveFirst
        If Not IsNull(rsLi!linr) Then
            lCouponLinr = rsLi!linr
        End If
    End If
    rsLi.Close: Set rsLi = Nothing
    
    
    If lCouponLinr = 0 Then
        anzeige "rot", "Es kann kein Coupon-Lieferant ermittelt werden.", Label1(4)
        Exit Sub
    End If
    
    loeschNEW "COUPONPRINT", gdBase
    CreateTableT2 "COUPONPRINT", gdBase
    
'    sSQL = "Insert into COUPONPRINT Select  "
'    sSQL = sSQL & " ARTNR  "
'    sSQL = sSQL & ", BEZEICH "
'    sSQL = sSQL & ", 'DRONOVA' as KETTE "
'    sSQL = sSQL & ", 'DRONOVA' as GRUPPE "
'    sSQL = sSQL & ", '" & cBudniKundNr & "' as KUNDNR "
'    sSQL = sSQL & ", ADATE  "
'    sSQL = sSQL & ", EAN "
'    sSQL = sSQL & ", MENGE  "
'    sSQL = sSQL & ", PREIS from Kassjour"
'    sSQL = sSQL & " where "
'    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
'    sSQL = sSQL & " and artnr in(select artnr from artlief where Linr = " & lCouponLinr & ")"
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into COUPONPRINT Select  "
    sSQL = sSQL & " ARTNR  "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", 'DRONOVA' as KETTE "
    sSQL = sSQL & ", 'DRONOVA' as GRUPPE "
    sSQL = sSQL & ", '" & cBudniKundNr & "' as KUNDNR "
    sSQL = sSQL & ", ADATE  "
    sSQL = sSQL & ", EAN "
    sSQL = sSQL & ", MENGE  "
    sSQL = sSQL & ", PREIS from Kassjour"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    sSQL = sSQL & " and ean like '98232*'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COUPONPRINT Set "
    sSQL = sSQL & " PREIS  =PREIS * (-1)  "
    sSQL = sSQL & ", von = " & CLng(cVon)
    sSQL = sSQL & ", bis = " & CLng(cBis)
    gdBase.Execute sSQL, dbFailOnError
    

    If Datendrin("COUPONPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKL45e"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
    
''''    If Datendrin("COUPONPRINT", gdBase) Then
''''        iRet = MsgBox("Möchten Sie diese Daten jetzt verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
''''        If iRet = vbYes Then
''''            schreibe_CouponCSV cBudniKundNr
''''
''''            giKissFtpMode = 43
''''            frmWKL38.Show 1
''''
''''
''''
''''        Else
'''''            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
''''        End If
''''        anzeige "normal", "", Label1(4)
''''    Else
''''        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
''''    End If
    
    
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenCouponEinloesungen"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenKedittilgung()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "KREDITZAPRINT", gdBase
    CreateTable "KREDITZAPRINT", gdBase
    
    sSQL = "Insert into KREDITZAPRINT Select "
    sSQL = sSQL & " ADATE "
    sSQL = sSQL & ", KREADATE "
    sSQL = sSQL & ", AZEIT "
    sSQL = sSQL & ", BEDNU "
    sSQL = sSQL & ", Menge "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", ARTNR "
    sSQL = sSQL & ", KK_ART "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", SENDOK from KREDITZA"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KREDITZAPRINT inner join Bedname on KREDITZAPRINT.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET KREDITZAPRINT.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KREDITZAPRINT inner join KUNDEN on KREDITZAPRINT.KUNDNR = KUNDEN.KUNDNR "
    sSQL = sSQL & " SET KREDITZAPRINT.NAME = KUNDEN.name "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KREDITZAPRINT inner join ARTIKEL on KREDITZAPRINT.ARTNR = ARTIKEL.ARTNR "
    sSQL = sSQL & " SET KREDITZAPRINT.BEZEICH = ARTIKEL.BEZEICH "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KREDITZAPRINT inner join Kassjour on KREDITZAPRINT.ARTNR = Kassjour.ARTNR "
    
    sSQL = sSQL & " and KREDITZAPRINT.MENGE = Kassjour.MENGE "
    sSQL = sSQL & " and KREDITZAPRINT.KUNDNR = Kassjour.KUNDNR "
    sSQL = sSQL & " and KREDITZAPRINT.BEDNU = Kassjour.BEDIENER "
    sSQL = sSQL & " and KREDITZAPRINT.adate = Kassjour.adate "
    sSQL = sSQL & " and KREDITZAPRINT.FILIALE = Kassjour.FILIALE "
    
    sSQL = sSQL & " SET KREDITZAPRINT.Preis = Kassjour.Preis "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("KREDITZAPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKL45c"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenKedittilgung"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenKunden_Auslieferung()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "KUNDAUSLIEFPRINT", gdBase
    CreateTable "KUNDAUSLIEFPRINT", gdBase
    
    sSQL = "Insert into KUNDAUSLIEFPRINT Select * from KUNDAUSLIEF"
    sSQL = sSQL & " where "
    sSQL = sSQL & " bestelltam >= " & cVon & " and bestelltam <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAUSLIEFPRINT inner join Bedname on KUNDAUSLIEFPRINT.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET KUNDAUSLIEFPRINT.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update KUNDAUSLIEFPRINT inner join KUNDEN on KUNDAUSLIEFPRINT.KUNDNR = KUNDEN.KUNDNR "
    sSQL = sSQL & " SET KUNDAUSLIEFPRINT.NAME = KUNDEN.name "
    gdBase.Execute sSQL, dbFailOnError

    
    If Datendrin("KUNDAUSLIEFPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

        reportbildschirm "", "aWKL136c"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenKunden_Auslieferung"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheDatenAboPlus()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "APPRINT", gdBase
    CreateTableT2 "APPRINT", gdBase
    
    sSQL = "Insert into APPRINT Select * from AboPlus_UMS"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("APPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

        reportbildschirm "", "aWKL45d"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
     
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenAboPlus"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExportiereAboPlus()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL                As String
    Dim cPfad               As String
    Dim cdatei              As String
    Dim cPfad1              As String
    Dim iRet                As Integer
    Dim rsrs                As Recordset
    Dim sAusgabedatname     As String
    Dim iFileNr             As Integer
    Dim lPos                As Long
    Dim cSatz               As String
    Dim cFeld               As String
    Dim dGeld               As Double
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = "Select * from APPRINT"
'    sSQL = sSQL & " where "
'    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        sAusgabedatname = "PVC_" & Format(DateValue(Now), "yyyymmdd") & ".txt"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
   
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = "978"
            cSatz = cSatz & "00000"
            
            If Not IsNull(rsrs!PFNR) Then
                cFeld = rsrs!PFNR
            End If
            
            While Len(Trim(cFeld)) < 8
                cFeld = "0" & cFeld
            Wend
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!ABOPLUSKARTE) Then
                cFeld = rsrs!ABOPLUSKARTE
            End If
            
            If Len(cFeld) = 13 Then
                cFeld = Mid(cFeld, 4, 9)
            Else
                cFeld = "000000000"
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!ADATE) Then
                cFeld = Format(rsrs!ADATE, "YYMMDD")
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!AZEIT) Then
                cFeld = Format(rsrs!AZEIT, "HHMM")
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!GELDWERT) Then
                dGeld = rsrs!GELDWERT
            End If
            cFeld = Format(dGeld, "0000000000.00")
            cFeld = SwapStr(cFeld, ",", "")
            cFeld = SwapStr(cFeld, "-", "")
            
'            dGeld = dGeld / 100
            
'            cFeld = Format(dGeld, "000000000000")
'
            While Len(Trim(cFeld)) < 12
                cFeld = "0" & cFeld
            Wend
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!BELEGNR) Then
                cFeld = rsrs!BELEGNR
            End If
            cSatz = cSatz & cFeld
            
            If dGeld >= 0 Then
                cSatz = cSatz & "00"
            Else
                cSatz = cSatz & "01"
            End If
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("APPRINT", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportiereAboPlus"
        Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim rec         As Recordset
    Dim cSatz       As String
    Dim cFeld       As String

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text1(0).Text = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
    Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
    Text1(1).Text = DateValue(Now) - 1
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDaten()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "EAPRINT", gdBase
    CreateTable "EAPRINT", gdBase
    
    sSQL = "Insert into EAPRINT Select *  from KAEINAUSF"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update EAPRINT inner join Bedname on EAPRINT.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET EAPRINT.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("EAPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKL136"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDaten"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDatenBedKasse()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "KASSBPRINT", gdBase
    CreateTable "KASSBPRINT", gdBase
    
    sSQL = "Insert into KASSBPRINT Select *  from KASSBEDP"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KASSBPRINT inner join Bedname on KASSBPRINT.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET KASSBPRINT.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("KASSBPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKL136b"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenBedKasse"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDaten2BedBeiStorno()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "STOPRINT", gdBase
    CreateTable "STOPRINT", gdBase
    
    sSQL = "Insert into STOPRINT Select *  from STORNO2"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update STOPRINT inner join Bedname on STOPRINT.bednu1 = Bedname.BEDNU "
    sSQL = sSQL & " SET STOPRINT.BEDNAME1 = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update STOPRINT inner join Bedname on STOPRINT.bednu2 = Bedname.BEDNU "
    sSQL = sSQL & " SET STOPRINT.BEDNAME2 = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("STOPRINT", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
        reportbildschirm "", "aWKL136a"
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Es sind keine Daten ermittelt worden.", Label1(4)
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDaten2BedBeiStorno"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "KASSBPRINT", gdBase
    loeschNEW "STOPRINT", gdBase
    loeschNEW "EAPRINT", gdBase
    loeschNEW "KREDITZAPRINT", gdBase
    loeschNEW "APPRINT", gdBase
    loeschNEW "KUNDAUSLIEFPRINT", gdBase
    loeschNEW "PLZPRINT", gdBase
    loeschNEW "PREISEDITKASSE_PRINT", gdBase
    
    loeschNEW "KKZAHLTE_PRINT", gdBase
    loeschNEW "ALTERG_PRINT", gdBase
    loeschNEW "COUPONPRINT", gdBase
    
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
Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 4     'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case 2
                        Text1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case Else
                        Text1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
        Case Is = 5     'ak monat
            Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
        Case Is = 6     'gestern
            Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now) - 1, "DD.MM.YY")
        Case Is = 7     'heute
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil allgemeine Kassenvorgänge ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

