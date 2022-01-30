VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL125 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kundenduplikatssuche"
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
   Begin VB.OptionButton Option7 
      Caption         =   "Übereinstimmung in Strasse, PLZ und Ort "
      Height          =   615
      Left            =   5520
      TabIndex        =   26
      Top             =   1680
      Width           =   3615
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Übereinstimmung in Name,Vorname,Strasse (die ersten 5 Buchstaben) und Ort "
      Height          =   615
      Left            =   5520
      TabIndex        =   25
      Top             =   960
      Width           =   3615
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   24
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
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   5
      Left            =   11400
      TabIndex        =   23
      Top             =   2880
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   ""
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   18
      Top             =   7440
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Protokolle"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Bonus übertragen"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kundenverkäufe übertragen"
      Height          =   375
      Left            =   9600
      TabIndex        =   16
      Top             =   3360
      Width           =   2055
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   13
      Top             =   6960
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Auflösen/Übertragen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   9960
      TabIndex        =   12
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      Left            =   9960
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Kundendaten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Übereinstimmung in Ortsname"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Value           =   -1  'True
      Width           =   4575
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Übereinstimmung in Plz"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Übereinstimmung in Name,Vorname und Plz"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Übereinstimmung in Geburtstag,Name und Vorname"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Übereinstimmung in Name,Vorname und Ort "
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4575
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   2
      Top             =   1080
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
      Caption         =   "Suche Daten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   1
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   12632256
      FocusRect       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Verkäufe"
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Verkäufe von insgesamt:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9600
      TabIndex        =   21
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Bonus"
      Height          =   255
      Index           =   3
      Left            =   9600
      TabIndex        =   20
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Bonus von:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9600
      TabIndex        =   19
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "und auf diesem übertragen"
      Height          =   495
      Index           =   1
      Left            =   9960
      TabIndex        =   15
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Dieses Kundenkonto wird aufgelöst"
      Height          =   855
      Index           =   0
      Left            =   9960
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblAnzeige 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   7800
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
      Caption         =   "Kundenduplikatssuche"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmWKL125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
        
            Dim iRet As Integer
            Dim sSQL As String
            
            If NewTableSuchenDBKombi("KUNDEDU", gdBase) Then
                If Datendrin("KUNDEDU", gdBase) Then
                    iRet = MsgBox("Möchten Sie die Kundenduplikate nocheinmal neu erstellen", vbYesNo + vbQuestion, "Zentrale Frage:")
                Else
                    iRet = vbYes
                End If
            Else
                iRet = vbYes
            End If
    
            If iRet = vbYes Then
                If Option1.Value = True Then
                    If SucheDaten Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option2.Value = True Then
                    If Suchedaten1 Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option3.Value = True Then
                    If SucheDaten2 Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option4.Value = True Then
                    If Suchedaten3 Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option5.Value = True Then
                    If Suchedaten4 Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option6.Value = True Then
                    If SucheDaten5 Then
                        Zusammenstellunganzeigen
                    End If
                ElseIf Option7.Value = True Then
                    If SucheDaten6 Then
                        Zusammenstellunganzeigen
                    End If
                End If
            Else
                Zusammenstellunganzeigen
            End If
        Case Is = 1
            voreinstellungspeichern
            Unload frmWKL125
        Case Is = 2
            UebertragOrAuflös
        Case 3
            
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "geloeschteKunden.rtf"
            Screen.MousePointer = 0
    
        Case 4
            If MSHFLEX1.Visible = True Then
                If MSHFLEX1.Row > 0 Then
                    gcKundenNr = MSHFLEX1.TextMatrix(MSHFLEX1.Row, 1)
                    anzeige "normal", "Die Kundendaten(" & gcKundenNr & ") werden angezeigt...", lblAnzeige
                    If gcKundenNr <> "" Then
                        Screen.MousePointer = 11
                        iKasse = 2
                        frmWKL13.Show 1
                        anzeige "normal", "", lblAnzeige
                        Screen.MousePointer = 0
                    End If
                Else
                    anzeige "rot", "Markieren Sie eine Zeile!", lblAnzeige
                End If
            End If
        Case 5
            If MSHFLEX1.Visible = True Then
                gckundnr = Text1(0).Text
                gckundnr = Trim$(gckundnr)
                gsARTNR = ""

                If gckundnr <> "" Then
                    frmWKL74.Show 1
                End If
            End If
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Long
    Dim i           As Long
    Dim j           As Long
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zusammenstellunganzeigen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zusammenstellunganzeigen()
    On Error GoTo LOKAL_ERROR
    
    Tabelleerstellen
    
    If NewTableSuchenDBKombi("KUNDEDU", gdBase) Then
        Tabellefuellen
        
'        FaerbeKU MSHFLEX1, 1, lblAnzeige
        
        Tabellenbreiteanpassen MSHFLEX1, 1 * gdTabfak
    Else
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblAnzeige
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zusammenstellunganzeigen"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoeschKunde(cKdnr As String)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim ctmp As String
    
    Dim rsrs As Recordset
            
    ctmp = "Kunde: " & cKdnr & " "
    ctmp = ctmp & WhatIsXfromKu(cKdnr, "Name") & " "
    ctmp = ctmp & WhatIsXfromKu(cKdnr, "VORName") & " wurde gelöscht."
    
    schreibeProtokollgKUN ctmp
    
    cSQL = "Update KUNDEN set synstatus = 'D',status = 'D' where KUNDNR = " & cKdnr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from KUNDEDU  where KUNDNR = " & cKdnr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from KULOESCH where KUNDNR = -1"
    Set rsrs = gdBase.OpenRecordset(cSQL)
        
    rsrs.AddNew
    
    rsrs!Adate = Format$(Now, "HH:MM:SS")
    rsrs!AZEIT = Fix(Now)
    rsrs!Kundnr = cKdnr
    rsrs!BEDNU = "99"
    rsrs!FILIALE = gbyteFilnr
    rsrs!SENDOK = False
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoeschKunde"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
 
    

End Sub
Private Sub UbertrageKunde(bBonus As Boolean, bVerkaufe As Boolean, cDelKdnr As String, cUKdnr As String)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim ctmp As String
    Dim sBonuD As String
    Dim sBonuU As String
    Dim dBoniges As Double
    Dim rsrs As Recordset
            
    ctmp = "Kunde: " & cDelKdnr & " "
    ctmp = ctmp & WhatIsXfromKu(cDelKdnr, "Name") & " "
    ctmp = ctmp & WhatIsXfromKu(cDelKdnr, "VORName") & " wurde übertragen auf : " & cUKdnr
    
    schreibeProtokollgKUN ctmp
    
    If bBonus = True Then
        sBonuD = ermBonusTotal(cDelKdnr)
        sBonuU = ermBonusTotal(cUKdnr)
        dBoniges = CDbl(sBonuD) + CDbl(sBonuU)
        cSQL = "Select Bonus,Status,Synstatus from kunden where kundnr = " & cUKdnr
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.Edit
            rsrs!Status = "E"
            rsrs!SYNStatus = "E"
            rsrs!BONUS = dBoniges
            rsrs.Update
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
    End If
    
    
    If bVerkaufe = True Then
    
        cSQL = "Update Kassjour set Kundnr  = " & cUKdnr & " "
        cSQL = cSQL & " Where kundnr = " & cDelKdnr
        gdBase.Execute cSQL, dbFailOnError
    Else
        cSQL = "Update Kassjour set Kundnr  = 0 "
        cSQL = cSQL & " Where kundnr = " & cDelKdnr
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UbertrageKunde"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub UebertragOrAuflös()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet        As Integer
    Dim bBonus      As Boolean
    Dim bVerkaufe   As Boolean
    Dim j           As Integer
    Dim lrow        As Long
    
    
    If Text1(0).Text = "" Then
        anzeige "rot", "Bitte eine Kundennummer eingeben!", lblAnzeige
        Text1(0).SetFocus
        Exit Sub
    
    Else
        If IsNumeric(Text1(0).Text) = True Then
            If Text1(1).Text = "" Then
                iRet = MsgBox("Möchten Sie wirklich das Kundenkonto " & Text1(0).Text & " löschen?", vbQuestion + vbYesNo, "Zentrale Frage:")
                If iRet = vbYes Then
                
                    lrow = MSHFLEX1.Row
                    LoeschKunde Trim(Text1(0).Text)
                    
                    Zusammenstellunganzeigen
                    
                    MSHFLEX1.Col = 1
                    If MSHFLEX1.Rows > lrow + 1 Then
                        MSHFLEX1.Row = lrow + 1
                        MSHFLEX1.TopRow = lrow + 1
                    
                    Else
                        MSHFLEX1.Row = MSHFLEX1.Rows - 1
                        MSHFLEX1.TopRow = MSHFLEX1.Rows - 1
                    End If
                    MSHFLEX1.SetFocus
                   
                    Text1(0).Text = ""
                    Text1(1).Text = ""
                Else
                    anzeige "normal", "Vorgang wurde abgebrochen", lblAnzeige
                    Exit Sub
                End If
            Else
                If IsNumeric(Text1(1).Text) = True Then
                
                    If Kundevorhanden(Text1(1).Text) Then
                    
                    Else
                        anzeige "rot", "Diese Kundennummer ist nicht vorhanden.", lblAnzeige
                        Text1(1).SetFocus
                        Exit Sub
                    
                    End If
                
                    iRet = MsgBox("Möchten Sie wirklich das Kundenkonto " & Text1(0).Text & " auf das Kundenkonto " & Text1(1).Text & " übertragen?", vbQuestion + vbYesNo, "Zentrale Frage:")
                    If iRet = vbYes Then
                    
                        If Check1.Value = vbChecked Then
                            bVerkaufe = True
                        Else
                            bVerkaufe = False
                        End If
                        
                        If Check2.Value = vbChecked Then
                            bBonus = True
                        Else
                            bBonus = False
                        End If
                        
                        If bVerkaufe = False And bBonus = False Then
                            iRet = MsgBox("Haben Sie alle zusätzlichen Optionen eingestellt?", vbQuestion + vbYesNo, "Zentrale Frage:")
                            If iRet = vbYes Then
                            
                            Else
                                anzeige "normal", "Vorgang wurde abgebrochen", lblAnzeige
                                Exit Sub
                            End If
                        
                        End If
                        
                        UbertrageKunde bBonus, bVerkaufe, Trim(Text1(0).Text), Trim(Text1(1).Text)
                        LoeschKunde Trim(Text1(0).Text)
                        
                        Zusammenstellunganzeigen
                        
                        For j = 2 To MSHFLEX1.Rows - 1
                            MSHFLEX1.Col = 1
                            MSHFLEX1.Row = j
                            If MSHFLEX1.Text = Trim(Text1(1).Text) Then
                                MSHFLEX1.Col = 1
                                If MSHFLEX1.Rows > j + 1 Then
                                    MSHFLEX1.Row = j + 1
                                    MSHFLEX1.TopRow = j + 1
                                Else
                                    MSHFLEX1.Row = MSHFLEX1.Rows - 1
                                    MSHFLEX1.TopRow = MSHFLEX1.Rows - 1
                                End If
                                MSHFLEX1.SetFocus
                                Exit For
                            End If
                        Next j
                        
                        Text1(0).Text = ""
                        Text1(1).Text = ""

                    Else
                        anzeige "normal", "Vorgang wurde abgebrochen", lblAnzeige
                        Exit Sub
                    End If
                
                
                Else
                    anzeige "rot", "Bitte eine gültige Kundennummer eingeben!", lblAnzeige
                    Text1(1).SetFocus
                    Exit Sub
                
                End If
            
            End If
        
        
        Else
            anzeige "rot", "Bitte eine gültige Kundennummer eingeben!", lblAnzeige
            Text1(0).SetFocus
            Exit Sub
        
        End If
    
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertragOrAuflös"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabelleerstellen()
    On Error GoTo LOKAL_ERROR

    
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 9
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 620
        .Text = "OK"
   
   
        
        .Col = 1
        .ColWidth(1) = 800
        .Text = "Kundennr"
        
        .Col = 2
        .ColWidth(2) = 1500
        .Text = "Vorname"
        
        .Col = 3
        .ColWidth(3) = 1600
        .Text = "Name"
        
        .Col = 4
        .ColWidth(4) = 1600
        .Text = "Straße"
        
        .Col = 5
        .ColWidth(5) = 600
        .Text = "Plz"
        
        .Col = 6
        .ColWidth(6) = 1600
        .Text = "Ort"
        
        .Col = 7
        .ColWidth(7) = 1200
        .Text = "Telefon"
        
        .Col = 8
        .ColWidth(8) = 1000
        .Text = "Geburtstag"
      

    
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabelleerstellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellefuellen()
    On Error GoTo LOKAL_ERROR

    Dim rsKUTEILME  As Recordset
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim lCounter    As Long
    Dim lMax        As Long
    
    
    Set rsKUTEILME = gdBase.OpenRecordset("KUNDEDU", dbOpenTable)
    
    lrow = 1
    If Not rsKUTEILME.EOF Then
        rsKUTEILME.MoveLast
        lMax = rsKUTEILME.RecordCount
        rsKUTEILME.MoveFirst
        
        MSHFLEX1.Redraw = False
        
        anzeige "normal", "Kunden werden ermittelt...", lblAnzeige
        
'        pbrZeit.Visible = True
'        pbrZeit.Max = 300
        
        Do While Not rsKUTEILME.EOF
            
            
            lMax = lMax - 1
            
            lrow = lrow + 1
            lCounter = lCounter + 1
            
            If lCounter = 300 Then
                anzeige "normal", "Kunden werden ermittelt(" & lMax & ")...", lblAnzeige
                lCounter = 0
            End If
'            pbrZeit.Value = lCounter
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = "X"
            
            If Not IsNull(rsKUTEILME!Kundnr) Then
                lWert = rsKUTEILME!Kundnr
            Else
                lWert = 0
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = lWert
            
            If Not IsNull(rsKUTEILME!vorname) Then
                sWert = rsKUTEILME!vorname
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!name) Then
                sWert = rsKUTEILME!name
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!strasse) Then
                sWert = rsKUTEILME!strasse
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 4
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!Plz) Then
                sWert = rsKUTEILME!Plz
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 5
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!Ort) Then
                sWert = rsKUTEILME!Ort
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 6
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!Tel) Then
                sWert = rsKUTEILME!Tel
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 7
            MSHFLEX1.Text = Trim$(sWert)
            
            If Not IsNull(rsKUTEILME!Datum1) Then
                sWert = rsKUTEILME!Datum1
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 8
            MSHFLEX1.Text = Trim$(sWert)
            
            rsKUTEILME.MoveNext
        Loop
'        pbrZeit.Visible = False
    End If
    rsKUTEILME.Close
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    If lrow > 1 Then
        anzeige "normal", lrow & " Kunden wurden ermittelt.", lblAnzeige
    
        
    ElseIf lrow = 1 Then
        anzeige "normal", lrow & " Kunde wurden ermittelt.", lblAnzeige
        
        
    Else
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblAnzeige
        
       
'        pbrZeit.Visible = False
        Exit Sub
    End If
    
'    fraZuErstellen.Visible = False
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellefuellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 11
        gsHelpstring = "Kundenduplikatssuche"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    If NewTableSuchenDBKombi("E93", gdApp) Then
        voreinstellungladen
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    loeschNEW "KUNDIDU1", gdBase
    loeschNEW "KDORT", gdBase
    loeschNEW "KUNDIDU", gdBase
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_SelChange()
    MSHFLEX1_Click
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 0, 1 'kundnr
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
   
    
    If KeyCode = vbKeyReturn Then
        Command1_Click 2
    End If
    
    If KeyCode = vbKeyEscape Then
        Command1_Click 1
    End If

   

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Set rsrs = gdApp.OpenRecordset("E93")
    
    If Not rsrs.EOF Then
        Option1.Value = rsrs!bo1
        Option2.Value = rsrs!bo2
        Option3.Value = rsrs!bo3
        Option4.Value = rsrs!bo4
        Option5.Value = rsrs!bo5
        Option6.Value = rsrs!bo6
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    
    loeschNEW "E93", gdApp
    CreateTable "E93", gdApp
    
    bo1 = Option1.Value
    bo2 = Option2.Value
    bo3 = Option3.Value
    bo4 = Option4.Value
    bo5 = Option5.Value
    bo6 = Option6.Value
   
    sSQL = "Insert into E93 (BO1,BO2,BO3,BO4,BO5,BO6) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & "," & bo5 & "," & bo6 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function SucheDaten() As Boolean
    On Error GoTo LOKAL_ERROR
    
    
    Dim lAnzdupli        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    
    SucheDaten = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU1", gdBase
    
    sSQL = "select * into KUNDIDU1 from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase
    
    
    sSQL = "select count(name) as count ,name,vorname,stadt into KUNDIDU from KUNDIDU1 group by name,vorname,stadt having count(name) > 1"
'    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
'        MsgBox lAnzdupli
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!name) Then
                sName = rsrs!name
                If Not IsNull(rsrs!vorname) Then
                    sVorname = rsrs!vorname
                    If Not IsNull(rsrs!STADT) Then
                        sStadt = rsrs!STADT

                        sSQL = "Insert into KUNDEDU select "
                        sSQL = sSQL & " TEL "
                        sSQL = sSQL & ", FAXNR "
                        sSQL = sSQL & ", EMAIL "
                        sSQL = sSQL & ", MOBILTEL "
                        sSQL = sSQL & ", VORNAME "
                        sSQL = sSQL & ", KUNDNR "
                        sSQL = sSQL & ", NAME "
                        sSQL = sSQL & ", STRASSE "
                        sSQL = sSQL & ", PLZ "
                        sSQL = sSQL & ", stadt as ort "
                        sSQL = sSQL & ", TITEL "
                        sSQL = sSQL & ", FIRMA "
                        sSQL = sSQL & " from Kunden"
                        sSQL = sSQL & " where  "
                        sSQL = sSQL & " name = '" & sName & "'"
                        sSQL = sSQL & " and  vorname = '" & sVorname & "'"
                        sSQL = sSQL & " and stadt = '" & sStadt & "'"
                        gdBase.Execute sSQL, dbFailOnError


                    End If
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing


    
    SucheDaten = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function SucheDaten5() As Boolean
    On Error GoTo LOKAL_ERROR
    
    
    Dim lAnzdupli        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    Dim sStrasse    As String
    
    SucheDaten5 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU1", gdBase
    
    sSQL = "select * into KUNDIDU1 from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase
    
    
    sSQL = "select count(name) as count ,name,vorname,stadt ,left(Strasse,5) as stras into KUNDIDU from KUNDIDU1 group by name,vorname,stadt ,left(Strasse,5)  having count(name) > 1"
'    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
'        MsgBox lAnzdupli
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!name) Then
                sName = rsrs!name
                If Not IsNull(rsrs!vorname) Then
                    sVorname = rsrs!vorname
                    If Not IsNull(rsrs!STADT) Then
                        sStadt = rsrs!STADT
                        
                        If Not IsNull(rsrs!Stras) Then
                            sStrasse = rsrs!Stras
    
                            sSQL = "Insert into KUNDEDU select "
                            sSQL = sSQL & " TEL "
                            sSQL = sSQL & ", FAXNR "
                            sSQL = sSQL & ", EMAIL "
                            sSQL = sSQL & ", MOBILTEL "
                            sSQL = sSQL & ", VORNAME "
                            sSQL = sSQL & ", KUNDNR "
                            sSQL = sSQL & ", NAME "
                            sSQL = sSQL & ", STRASSE "
                            sSQL = sSQL & ", PLZ "
                            sSQL = sSQL & ", stadt as ort "
                            sSQL = sSQL & ", TITEL "
                            sSQL = sSQL & ", FIRMA "
                            sSQL = sSQL & " from Kunden"
                            sSQL = sSQL & " where  "
                            sSQL = sSQL & " name = '" & sName & "'"
                            sSQL = sSQL & " and  vorname = '" & sVorname & "'"
                            sSQL = sSQL & " and stadt = '" & sStadt & "'"
                            sSQL = sSQL & " and strasse like '" & sStrasse & "*'"
                            gdBase.Execute sSQL, dbFailOnError
                        End If
                    End If
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    SucheDaten5 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten5"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SucheDaten2() As Boolean
    On Error GoTo LOKAL_ERROR

    Dim lAnzdupli   As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    
    SucheDaten2 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU1", gdBase
    
    sSQL = "select * into KUNDIDU1 from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase
    
    
    sSQL = "select count(name) as count ,name,vorname,plz into KUNDIDU from KUNDIDU1 group by name,vorname,plz having count(name) > 1"
'    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
'        MsgBox lAnzdupli
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!name) Then
                sName = rsrs!name
                If Not IsNull(rsrs!vorname) Then
                    sVorname = rsrs!vorname
                    If Not IsNull(rsrs!Plz) Then
                        sStadt = rsrs!Plz

                        sSQL = "Insert into KUNDEDU select "
                        sSQL = sSQL & " TEL "
                        sSQL = sSQL & ", FAXNR "
                        sSQL = sSQL & ", EMAIL "
                        sSQL = sSQL & ", MOBILTEL "
                        sSQL = sSQL & ", VORNAME "
                        sSQL = sSQL & ", KUNDNR "
                        sSQL = sSQL & ", NAME "
                        sSQL = sSQL & ", STRASSE "
                        sSQL = sSQL & ", PLZ "
                        sSQL = sSQL & ", stadt as ort "
                        sSQL = sSQL & ", TITEL "
                        sSQL = sSQL & ", FIRMA "
                        sSQL = sSQL & " from Kunden"
                        sSQL = sSQL & " where  "
                        sSQL = sSQL & " name = '" & sName & "'"
                        sSQL = sSQL & " and  vorname = '" & sVorname & "'"
                        sSQL = sSQL & " and plz = '" & sStadt & "'"
                        gdBase.Execute sSQL, dbFailOnError


                    End If
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing


    
    SucheDaten2 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten2"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function Suchedaten3() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzdupli        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    Dim sPlz        As String
    
    Suchedaten3 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase
    
    loeschNEW "KDORT", gdBase
    sSQL = "select stadt ,plz,status into KDORT from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    sSQL = "select count(stadt) as count,plz into KUNDIDU from KDORT group by stadt,plz having count(stadt) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KDPLZ", gdBase
    
    sSQL = "select count(plz) as count,plz into KDPLZ from KUNDIDU group by plz having count(plz) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KDPLZ", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
    End If

    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Ortsbezeichnungen werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KDPLZ")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!Plz) Then
                sPlz = rsrs!Plz
                
                sSQL = "Insert into KUNDEDU select distinct(stadt) as ort,plz "
                sSQL = sSQL & " from Kunden"
                sSQL = sSQL & " where  "
                sSQL = sSQL & " plz = '" & sPlz & "'"
                gdBase.Execute sSQL, dbFailOnError

            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing


    
    Suchedaten3 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten3"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function Suchedaten4() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzdupli        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    Dim sPlz        As String
    
    Suchedaten4 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    
    
    loeschNEW "KUNDIDU", gdBase
    
    loeschNEW "KDORT", gdBase
    sSQL = "select stadt ,plz into KDORT from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    
    sSQL = "select count(plz) as count,stadt into KUNDIDU from KDORT group by plz,stadt having count(plz) > 1"
    
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KDPLZ", gdBase
    
    sSQL = "select count(stadt) as count,stadt into KDPLZ from KUNDIDU group by stadt having count(stadt) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KDPLZ", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
    End If

    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Postleitzahlen werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KDPLZ")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!STADT) Then
                sStadt = rsrs!STADT
                
                sSQL = "Insert into KUNDEDU select distinct(Plz) as name ,stadt as ort "
                sSQL = sSQL & " from Kunden"
                sSQL = sSQL & " where  "
                sSQL = sSQL & " stadt = '" & sStadt & "'"
                gdBase.Execute sSQL, dbFailOnError

            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    sSQL = "Update KUNDEDU set Plz = name  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDEDU set name = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    Suchedaten4 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten4"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function Suchedaten1() As Boolean
    On Error GoTo LOKAL_ERROR
    
    
    Dim lAnzdupli        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sName       As String
    Dim sVorname    As String
    Dim sStadt      As String
    
    Suchedaten1 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU1", gdBase
    
    sSQL = "select * into KUNDIDU1 from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase


    sSQL = "select count(datum1) as count ,datum1,name,vorname into KUNDIDU from KUNDIDU1 group by name,vorname, datum1 having count(datum1) > 1"
'    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
'        MsgBox lAnzdupli
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!name) Then
                sName = rsrs!name
                If Not IsNull(rsrs!vorname) Then
                    sVorname = rsrs!vorname
                    If Not IsNull(rsrs!Datum1) Then
                        sStadt = rsrs!Datum1
                        
                        If sStadt <> "00:00:00" Then
                            If sName <> "" Then
                                If sVorname <> "" Then
                                    sSQL = "Insert into KUNDEDU select "
                                    sSQL = sSQL & " TEL "
                                    sSQL = sSQL & ", FAXNR "
                                    sSQL = sSQL & ", EMAIL "
                                    sSQL = sSQL & ", MOBILTEL "
                                    sSQL = sSQL & ", VORNAME "
                                    sSQL = sSQL & ", KUNDNR "
                                    sSQL = sSQL & ", NAME "
                                    sSQL = sSQL & ", STRASSE "
                                    sSQL = sSQL & ", PLZ "
                                    sSQL = sSQL & ", stadt as ort "
                                    sSQL = sSQL & ", TITEL "
                                    sSQL = sSQL & ", FIRMA "
                                    sSQL = sSQL & ", datum1 "
                                    sSQL = sSQL & " from Kunden"
                                    sSQL = sSQL & " where  "
                                    sSQL = sSQL & " name = '" & sName & "'"
                                    sSQL = sSQL & " and  vorname = '" & sVorname & "'"
                                    sSQL = sSQL & " and datum1 = " & CLng(DateValue(sStadt))
                                    gdBase.Execute sSQL, dbFailOnError
                                End If
                            End If
                        End If

                    End If
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Suchedaten1 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten1"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function SucheDaten6() As Boolean
    On Error GoTo LOKAL_ERROR

    Dim lAnzdupli   As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsKuD       As Recordset
    Dim sStrasse    As String
    Dim sPlz        As String
    Dim sStadt      As String
    
    SucheDaten6 = False
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDIDU1", gdBase
    
    sSQL = "select * into KUNDIDU1 from Kunden "
    sSQL = sSQL & " where (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden werden ermittelt, bitte warten........................", lblAnzeige
    
    loeschNEW "KUNDIDU", gdBase
    
    
    sSQL = "select count(name) as count ,strasse,stadt,plz into KUNDIDU from KUNDIDU1 group by strasse,stadt,plz having count(strasse) > 1"
'    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
'        MsgBox lAnzdupli
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
    
    loeschNEW "KUNDEDU", gdBase
    CreateTable "KUNDEDU", gdBase
    
    Set rsrs = gdBase.OpenRecordset("KUNDIDU")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAnzdupli = lAnzdupli - 1
            anzeige "normal", lAnzdupli & " verschiedene Kunden werden angezeigt, bitte warten...", lblAnzeige
            If Not IsNull(rsrs!strasse) Then
                sStrasse = rsrs!strasse
                If Not IsNull(rsrs!STADT) Then
                    sStadt = rsrs!STADT
                    If Not IsNull(rsrs!Plz) Then
                        sPlz = rsrs!Plz

                        sSQL = "Insert into KUNDEDU select "
                        sSQL = sSQL & " TEL "
                        sSQL = sSQL & ", FAXNR "
                        sSQL = sSQL & ", EMAIL "
                        sSQL = sSQL & ", MOBILTEL "
                        sSQL = sSQL & ", VORNAME "
                        sSQL = sSQL & ", KUNDNR "
                        sSQL = sSQL & ", NAME "
                        sSQL = sSQL & ", STRASSE "
                        sSQL = sSQL & ", PLZ "
                        sSQL = sSQL & ", stadt as ort "
                        sSQL = sSQL & ", TITEL "
                        sSQL = sSQL & ", FIRMA "
                        sSQL = sSQL & " from Kunden"
                        sSQL = sSQL & " where  "
                        sSQL = sSQL & " strasse = '" & sStrasse & "'"
                        sSQL = sSQL & " and  stadt = '" & sStadt & "'"
                        sSQL = sSQL & " and plz = '" & sPlz & "'"
                        gdBase.Execute sSQL, dbFailOnError


                    End If
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing


    
    SucheDaten6 = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten6"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Sub MSHFLEX1_Click()
On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row > 0 Then
        Text1(0).Text = MSHFLEX1.TextMatrix(MSHFLEX1.Row, 1)
        Label1(2).Caption = "Bonus von: " & Text1(0).Text
        Label1(3).Caption = ermBonusTotal(Text1(0).Text)
        Label1(4).Caption = "Verkäufe von: " & Text1(0).Text & " insgesamt "
        Label1(5).Caption = LeseVerkäufeKundeTotal(Text1(0).Text)
    End If
    
Exit Sub
    
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenduplikatssuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row > 1 Then
        
    Else
        sortierenHGrid MSHFLEX1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

