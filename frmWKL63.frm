VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL63 
   Caption         =   "Artikel Verkauf"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL63.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   0
      Left            =   9720
      TabIndex        =   3
      Top             =   7920
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11775
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
         Height          =   1320
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "mit Kollegenverkäufen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         TabIndex        =   20
         Top             =   6480
         Value           =   1  'Aktiviert
         Width           =   2655
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
         Height          =   1320
         Left            =   3480
         TabIndex        =   16
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   1935
         Left            =   9600
         TabIndex        =   6
         Top             =   840
         Width           =   2175
         Begin VB.OptionButton Option4 
            Caption         =   "Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Tag             =   "adate desc"
            Top             =   360
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Bediener"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Tag             =   "Bediener"
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Tag             =   "Filiale"
            Top             =   840
            Width           =   3255
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Kunden Nr."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   9
            Tag             =   "kundnr desc"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Menge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   8
            Tag             =   "menge desc"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   7
            Tag             =   "Filiale , adate desc"
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Sortierung nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4620
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9375
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Jahressummen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Ø Nettospannen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Ø erzielte Nettospanne"
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
         Left            =   6720
         TabIndex        =   18
         Top             =   5640
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
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
         Left            =   10080
         TabIndex        =   17
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "insgesamt verk:"
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
         Index           =   0
         Left            =   9600
         TabIndex        =   15
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Artikelanzahl"
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
         Index           =   9
         Left            =   9600
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   12600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel Verkauf"
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
      TabIndex        =   5
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label lblanzeige 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL63"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim siENS As Single

Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    With Frame5
        .Height = 6855
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR
    
    If Check1.Value = vbChecked Then
        If Option4(3).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(3).Tag, True
        ElseIf Option4(4).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(4).Tag, True
        ElseIf Option4(5).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(5).Tag, True
        ElseIf Option4(6).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(6).Tag, True
        ElseIf Option4(7).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(7).Tag, True
        ElseIf Option4(8).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(8).Tag, True
        End If
            
    Else
    
        If Option4(3).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(3).Tag, False
        ElseIf Option4(4).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(4).Tag, False
        ElseIf Option4(5).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(5).Tag, False
        ElseIf Option4(6).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(6).Tag, False
        ElseIf Option4(7).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(7).Tag, False
        ElseIf Option4(8).Value = True Then
            ZeigeVerkäufe List3, gsARTNR, Option4(8).Tag, False
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 0
        Unload frmWKL63
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren

'    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    List1.AddItem "Datum     Uhrzeit Menge  Fil  Preis   KundNr    Name          Bed  NSP"
    
    If gsARTNR = "" Then
        Exit Sub
    End If

    ZeigeVerkäufe List3, gsARTNR, "adate desc", True
    ZeigeNSMITTELWERTE gsARTNR
    ZeigeSummen gsARTNR
    anzeige "normal", gsARTNR, lblanzeige
    
    Label1(9).Caption = ermMENGE(gsARTNR, "", "", 0)
    Label1(9).Refresh
    
    Label2.Caption = Format(siENS, "##0.00")
    Label2.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub ZeigeVerkäufe(Listx As ListBox, sarti As String, sOrder As String, bmitKOLLVK)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim sSatz       As String
    Dim rsrs        As Recordset
    Dim ENS         As Single
    Dim cMWST       As String
    Dim ENSMITTEL   As Single
    Dim lcount      As Long
    Dim dEkpr       As Double
    lcount = 0
    
    Screen.MousePointer = 11
    List3.Clear
    List3.Visible = False
    
    loeschNEW "ARTANZ" & srechnertab, gdBase
    CreateTableT2 "ARTANZ" & srechnertab, gdBase
    
    sSQL = "Insert into ARTANZ" & srechnertab
    sSQL = sSQL & " select "
    sSQL = sSQL & " MENGE  "
    sSQL = sSQL & " , PREIS  "
    sSQL = sSQL & " , MWST  "
    sSQL = sSQL & " , BEDIENER  "
    sSQL = sSQL & " , EKPR  "
    sSQL = sSQL & " , KUNDNR "
    sSQL = sSQL & " , FILIALE  "
    sSQL = sSQL & " , ADATE  "
    sSQL = sSQL & " , AZEIT  "
    sSQL = sSQL & " , 'Kassjour' as  QTAB "
    sSQL = sSQL & " from Kassjour where artnr = " & sarti
    gdBase.Execute sSQL, dbFailOnError
    
    If bmitKOLLVK Then
        sSQL = "Insert into ARTANZ" & srechnertab
        sSQL = sSQL & " select "
        sSQL = sSQL & " MENGE  "
        sSQL = sSQL & " , PREIS  "
        sSQL = sSQL & " , MWST  "
        sSQL = sSQL & " , BEDIENER  "
        sSQL = sSQL & " , EKPR  "
        sSQL = sSQL & " , KUNDNR "
        sSQL = sSQL & " , FILIALE  "
        sSQL = sSQL & " , ADATE  "
        sSQL = sSQL & " , AZEIT  "
        sSQL = sSQL & " , 'Kollverk' as  QTAB "
        sSQL = sSQL & " from Kollverk where artnr = " & sarti
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    
    
    sSQL = "Update ARTANZ" & srechnertab & " inner join KUNDEN on ARTANZ" & srechnertab & ".KUNDNR = KUNDEN.KUNDNR"
    sSQL = sSQL & " SET ARTANZ" & srechnertab & ".KUNAME = KUNDEN.NAME "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Update ARTANZ" & srechnertab
    sSQL = sSQL & " set Preis = 0 where Preis is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select * from ARTANZ" & srechnertab
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ADATE) Then
            sSatz = Format(rsrs!ADATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!ADATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!AZEIT) Then
            sSatz = sSatz & Format(rsrs!AZEIT, "HH:MM") & Space(12 - Len(Format(rsrs!AZEIT, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!Menge) Then
            sSatz = sSatz & rsrs!Menge & Space(4 - Len(rsrs!Menge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(4 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(4)
        End If
        
        If Not IsNull(rsrs!Preis) Then
            sSatz = sSatz & Format$(rsrs!Preis, "####0.00") & Space(8 - Len(Format$(rsrs!Preis, "####0.00")))
        Else
            sSatz = sSatz & Space(8)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(10 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(10)
        End If
        
        If Not IsNull(rsrs!KUNAME) Then
            If Len(rsrs!KUNAME) > 13 Then
                sSatz = sSatz & Left(rsrs!KUNAME, 10) & "..."
            Else
                sSatz = sSatz & rsrs!KUNAME & Space(13 - Len(rsrs!KUNAME))
            End If
        Else
            sSatz = sSatz & Space(13)
        End If
        
        sSatz = sSatz & Space(1)
        
        If Not IsNull(rsrs!ekpr) Then
            dEkpr = rsrs!ekpr
        Else
            dEkpr = 0
        End If
        
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER & Space(5 - Len(rsrs!BEDIENER))
        Else
            sSatz = sSatz & Space(5)
        End If
        
        ENS = 0
        If Not IsNull(rsrs!MWST) Then
            cMWST = rsrs!MWST
            
            If CDbl(rsrs!Preis) > 0# Then
                Select Case cMWST
                    Case Is = "V"
                        ENS = ((((rsrs!Preis / (100 + gdMWStV)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStV)) * 100)
                    Case Is = "E"
                    
                        ENS = ((((rsrs!Preis / (100 + gdMWStE)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStE)) * 100)
                    Case Is = "O"
                        ENS = ((((rsrs!Preis / (100 + gdMWStO)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStO)) * 100)
                End Select
            ElseIf CDbl(rsrs!Preis) < 0 Then
                Select Case cMWST
                    Case Is = "V"
                        ENS = ((((rsrs!Preis / (100 + gdMWStV)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStV)) * 100)
                    Case Is = "E"
                    
                        ENS = ((((rsrs!Preis / (100 + gdMWStE)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStE)) * 100)
                    Case Is = "O"
                        ENS = ((((rsrs!Preis / (100 + gdMWStO)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStO)) * 100)
                End Select
                
                ENS = ENS * (-1)
            
            End If
            
            ENSMITTEL = ENSMITTEL + ENS
            lcount = lcount + 1
             
            sSatz = sSatz & Format$(ENS, "##0.00")
        End If

        List3.AddItem sSatz
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If lcount > 0 Then
        ENSMITTEL = ENSMITTEL / lcount
    Else
        ENSMITTEL = 0#
    End If
    
    List3.Visible = True
    
    siENS = ENSMITTEL
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeVerkäufe"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
'    Resume Next
End Sub

Private Sub ZeigeNSMITTELWERTE(sarti As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim sSatz       As String
    Dim rsrs        As Recordset
    Dim ENS         As Single
    Dim cMWST       As String
    Dim ENSMITTEL   As Single
    Dim iAktYEAR    As Integer
    Dim lcount      As Long
    Dim dEkpr       As Double
    lcount = 0
    iAktYEAR = 0
    Screen.MousePointer = 11
    List2.Clear

    sSQL = "select * from ARTANZ" & srechnertab
    sSQL = sSQL & " order by year(adate) desc "

'    sSQL = "select * from Kassjour where artnr = " & sarti
'    sSQL = sSQL & " order by year(adate) desc "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ekpr) Then
            dEkpr = rsrs!ekpr
        Else
            dEkpr = 0
        End If
        
        ENS = 0
        If Not IsNull(rsrs!MWST) Then
            cMWST = rsrs!MWST
            
            
            
            If CDbl(rsrs!Preis) > 0# Then
                Select Case cMWST
                    Case Is = "V"
                        ENS = ((((rsrs!Preis / (100 + gdMWStV)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStV)) * 100)
                    Case Is = "E"
                    
                        ENS = ((((rsrs!Preis / (100 + gdMWStE)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStE)) * 100)
                    Case Is = "O"
                        ENS = ((((rsrs!Preis / (100 + gdMWStO)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStO)) * 100)
                End Select
            ElseIf CDbl(rsrs!Preis) < 0 Then
                Select Case cMWST
                    Case Is = "V"
                        ENS = ((((rsrs!Preis / (100 + gdMWStV)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStV)) * 100)
                    Case Is = "E"
                    
                        ENS = ((((rsrs!Preis / (100 + gdMWStE)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStE)) * 100)
                    Case Is = "O"
                        ENS = ((((rsrs!Preis / (100 + gdMWStO)) * 100) - (dEkpr * rsrs!Menge)) * 100) / ((rsrs!Preis / (100 + gdMWStO)) * 100)
                End Select
                ENS = ENS * (-1)
            End If
            
            If Not IsNull(rsrs!ADATE) Then
            
                If iAktYEAR = Year(rsrs!ADATE) Then
                    ENSMITTEL = ENSMITTEL + ENS
                    
                    lcount = lcount + 1
                    
                Else
                    If iAktYEAR > 0 Then
                        If lcount > 0 Then
                            ENSMITTEL = ENSMITTEL / lcount
                        Else
                            ENSMITTEL = 0#
                        End If
                        List2.AddItem iAktYEAR & " " & Format(ENSMITTEL, "##0.00")
                    End If
                    
                    lcount = 1
                    ENSMITTEL = ENS
                    iAktYEAR = Year(rsrs!ADATE)
                End If
               
            End If
        End If

        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If iAktYEAR > 0 Then
        If lcount > 0 Then
            ENSMITTEL = ENSMITTEL / lcount
        Else
            ENSMITTEL = 0#
        End If
        List2.AddItem iAktYEAR & " " & Format(ENSMITTEL, "##0.00")
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeNSMITTELWERTE"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub ZeigeSummen(cArtNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sSatz As String
    
    List4.Clear
    sSQL = "select distinct(year(adate)) as jahr,sum(preis) as sumPreis,sum(menge) as sumMenge from Kassjour where ARTNR = " & cArtNr
    sSQL = sSQL & " group by year(adate) order by year(adate) desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!jahr) Then
            sSatz = rsrs!jahr & Space(4 - Len(rsrs!jahr))
        End If
        
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(9 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!sumMenge) Then
            sSatz = sSatz & Space(6 - Len(Format$(rsrs!sumMenge, "####0"))) & Format$(rsrs!sumMenge, "####0")
        Else
            sSatz = sSatz & Space(6)
        End If
        
        List4.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeSummen"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "ARTANZ" & srechnertab, gdBase
    
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
Private Sub Option4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Check1.Value = vbChecked Then
        ZeigeVerkäufe List3, gsARTNR, Option4(Index).Tag, True
    Else
        ZeigeVerkäufe List3, gsARTNR, Option4(Index).Tag, False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
