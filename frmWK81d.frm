VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWK81d 
   BackColor       =   &H00C0C000&
   Caption         =   "Termine - Feiertage"
   ClientHeight    =   8610
   ClientLeft      =   1995
   ClientTop       =   750
   ClientWidth     =   11910
   Icon            =   "frmWK81d.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11655
      Begin VB.CheckBox Check2 
         Caption         =   "beweglicher Feiertag"
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
         Left            =   6000
         TabIndex        =   12
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "bundesweit"
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
         Left            =   6000
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   2
         Top             =   960
         Width           =   1935
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   3735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
         Height          =   5055
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   8916
         _Version        =   393216
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
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   8
         Left            =   9600
         TabIndex        =   9
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
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
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   3
         Left            =   1320
         TabIndex        =   13
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum"
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
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Feiertag"
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
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   9720
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Termine - Feiertage"
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
      TabIndex        =   10
      Top             =   120
      Width           =   7455
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
End
Attribute VB_Name = "frmWK81d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerFeiertag As Byte
Dim SpaltennummerDatum As Byte
Dim SpaltennummerANWENDEN As Byte
Private Function fnPruefeEingabeWK81d() As Integer
    On Error GoTo LOKAL_ERROR
    
    fnPruefeEingabeWK81d = 0
    
    If Trim$(Text1(2).Text) = "" Then
        fnPruefeEingabeWK81d = 1
        Exit Function
    End If
    
    If Trim$(Text1(0).Text) = "" Then
        fnPruefeEingabeWK81d = 2
        Exit Function
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWK81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Zeige_Daten_81d()
On Error GoTo LOKAL_ERROR

    Tabcheck "FEIERTAGE"
    FormatGridOverTablay "FEIERTAGE"

    Dim j As Integer

    With MSHFLEX1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = Len(.TextMatrix(0, j)) * 80
        Next j
    End With

    FuellenMShFlex1WKL81d
    ermittlespalten
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_Daten_81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKL81d()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim sSQL        As String
    Dim iJahr       As Integer
    
    loeschNEW "FEIERTAGE_TEMP", gdBase
    CreateTableT2 "FEIERTAGE_TEMP", gdBase
    
    'die beweglichen
    sSQL = " Insert into FEIERTAGE_TEMP select "
    sSQL = sSQL & " FDAT & FDATJAHR as Datum  "
    sSQL = sSQL & ", FEIERTAGBEZ "
    sSQL = sSQL & ", BUNDESWEIT "
    sSQL = sSQL & ", ANWENDEN "
    sSQL = sSQL & " from Feiertage "
    sSQL = sSQL & " where  FDATJAHR <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'die fixen 1.Mai usw
    
    sSQL = " Insert into FEIERTAGE_TEMP select "
    sSQL = sSQL & " FDAT & Year(Now) as Datum  "
    sSQL = sSQL & ", FEIERTAGBEZ "
    sSQL = sSQL & ", BUNDESWEIT "
    sSQL = sSQL & ", ANWENDEN "
    sSQL = sSQL & " from Feiertage "
    sSQL = sSQL & " where  FDATJAHR = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into FEIERTAGE_TEMP select "
    sSQL = sSQL & " FDAT & Year(Now + 365) as Datum  "
    sSQL = sSQL & ", FEIERTAGBEZ "
    sSQL = sSQL & ", BUNDESWEIT "
    sSQL = sSQL & ", ANWENDEN "
    sSQL = sSQL & " from Feiertage "
    sSQL = sSQL & " where  FDATJAHR = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into FEIERTAGE_TEMP select "
    sSQL = sSQL & " FDAT & Year(Now + 722) as Datum  "
    sSQL = sSQL & ", FEIERTAGBEZ "
    sSQL = sSQL & ", BUNDESWEIT "
    sSQL = sSQL & ", ANWENDEN "
    sSQL = sSQL & " from Feiertage "
    sSQL = sSQL & " where  FDATJAHR = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into FEIERTAGE_TEMP select "
    sSQL = sSQL & " FDAT & Year(Now + 1095) as Datum  "
    sSQL = sSQL & ", FEIERTAGBEZ "
    sSQL = sSQL & ", BUNDESWEIT "
    sSQL = sSQL & ", ANWENDEN "
    sSQL = sSQL & " from Feiertage "
    sSQL = sSQL & " where  FDATJAHR = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Select * from FEIERTAGE_TEMP where datum >= datevalue(now) order by Datum asc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    MSHFLEX1.Redraw = False
    MSHFLEX1.Visible = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        

            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Row = 0
                MSHFLEX1.Col = i
                
                If sSpaltenname(i) = MSHFLEX1.Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "bundesweit"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            MSHFLEX1.Row = lrow
                            
                            If UCase(sWert) = "WAHR" Then
                                MSHFLEX1.Text = "bundesweit"
                            Else
                                MSHFLEX1.Text = ""
                            End If
                            
                        Case Is = "anwenden"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            MSHFLEX1.Row = lrow
                            
                            If UCase(sWert) = "WAHR" Then
                                MSHFLEX1.Text = "anwenden"
                                MSHFLEX1.CellFontBold = True
                                MSHFLEX1.CellForeColor = vbGreen
                                
                            Else
                                MSHFLEX1.Text = "nicht anwenden"
                                MSHFLEX1.CellFontBold = True
                                MSHFLEX1.CellForeColor = vbRed
                                
                                
                            End If

                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = sWert
                        
                    End Select
                    
            
                    If Len(MSHFLEX1.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                        aBreite(i) = Len(MSHFLEX1.TextMatrix(lrow, i)) * 80
                    End If
                    
                End If
            Next i
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        MSHFLEX1.Col = i
        MSHFLEX1.ColWidth(i) = aBreite(i) * 1.8
    Next i
        
    
    rsrs.Close: Set rsrs = Nothing
    
    If byAnzahlSpalten < 2 Then
    
    Else
        MSHFLEX1.FixedCols = 1
    End If
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    Screen.MousePointer = 0
        
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleDatenWK81d(sFeiert As String, sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Text1(2).Text = sdat
    Text1(0).Text = sFeiert
    
    cSQL = "Select * from Feiertage where FEIERTAGBEZ = '" & sFeiert & "' "
    cSQL = cSQL & " and FDAT = '" & Left(sdat, 6) & "'"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        Check2.Value = vbUnchecked
        If Not IsNull(rsrs!FDATJAHR) Then
            If rsrs!FDATJAHR <> "" Then
                Check2.Value = vbChecked
            End If
        End If
        
        Check1.Value = vbUnchecked
        If Not IsNull(rsrs!bundesweit) Then
            If rsrs!bundesweit = True Then
                Check1.Value = vbChecked
            End If
        End If
    Else
        Text1(2).Text = ""
        Text1(0).Text = ""
        Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Text1(0).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDatenWK81a"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheDatenWK81d(sFeiert As String, sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from Feiertage where FEIERTAGBEZ = '" & sFeiert & "' "
    cSQL = cSQL & " and FDAT = '" & Left(sdat, 6) & "'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.delete
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Zeige_Daten_81d
    LeereFelderWK81d

    Text1(0).SetFocus

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheDatenWK81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub drucke_Behandlungstexte()
'    On Error GoTo LOKAL_ERROR
'
'    Dim sSQL As String
'
'    loeschNEW "PRINT_BEHTEXTE", gdBase
'    CreateTableT2 "PRINT_BEHTEXTE", gdBase
'
'    sSQL = "Insert into PRINT_BEHTEXTE Select * from TERM_STD order by gliederung, bezeich"
'    gdBase.Execute sSQL, dbFailOnError
'
'    reportbildschirm "", "aWKL81a"
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "drucke_Behandlungstexte"
'    Fehler.gsFehlertext = "Im Programmteil Vorgaben Behandlungen ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Function ermMaxNr_Term() As Long
'    On Error GoTo LOKAL_ERROR
'
'    Dim cSQL As String
'    Dim rsRS As Recordset
'
'    ermMaxNr_Term = 0
'
'    cSQL = "Select max(nr) as maxi from TERM_STD "
'    Set rsRS = gdBase.OpenRecordset(cSQL)
'    If Not rsRS.EOF Then
'        If Not IsNull(rsRS!maxi) Then
'            ermMaxNr_Term = rsRS!maxi
'        End If
'    End If
'    rsRS.Close: Set rsRS = Nothing
'
'    ermMaxNr_Term = ermMaxNr_Term + 1
'
'Exit Function
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "ermMaxNr_Term"
'    Fehler.gsFehlertext = "Im Programmteil Vorgaben Behandlungen ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Function
Private Sub LeereFelderWK81d()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(2).Text = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereFelderWK81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWK81d()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cBezeich As String
    Dim cDatum As String
    
    cDatum = Text1(2).Text
    cBezeich = Text1(0).Text
    
    cSQL = "Select * from Feiertage where FEIERTAGBEZ = '" & cBezeich & "' "
    cSQL = cSQL & " and FDAT = '" & Left(cDatum, 6) & "'"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    rsrs!FEIERTAGBEZ = cBezeich
    If Check1.Value = vbChecked Then
        rsrs!bundesweit = True
    Else
        rsrs!bundesweit = False
    End If
    
    If Check2.Value = vbUnchecked Then
        rsrs!FDATJAHR = ""
    Else
        rsrs!FDATJAHR = Right(cDatum, 4)
    End If
    
    rsrs!FDAT = Left(cDatum, 6)
    rsrs!anwenden = True
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    Zeige_Daten_81d
    LeereFelderWK81d

    Text1(0).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWK81d"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sFeiertag As String
    Dim sDatum As String
    
    Select Case Index
        Case Is = 0     'Speichern
            iRet = fnPruefeEingabeWK81d()
            Select Case iRet
                Case Is = 0     'alles okay
                    SchreibeDatenWK81d
                    
                Case Is = 1     'keine Datum
                    MsgBox "Bitte ein Datum angeben!", vbInformation, "Winkiss Hinweis:"
                    Text1(2).SetFocus
                Case Is = 2     'keine Bez
                    MsgBox "Bitte eine Feiertagsbezeichnung angeben!", vbInformation, "Winkiss Hinweis:"
                    Text1(0).SetFocus
                
            End Select
        Case Is = 1     'Auswählen
            If MSHFLEX1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            sFeiertag = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerFeiertag))
            sDatum = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerDatum))

            HoleDatenWK81d sFeiertag, sDatum
            
        Case Is = 2     'Löschen
            If MSHFLEX1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            sFeiertag = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerFeiertag))
            sDatum = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerDatum))
            
            LoescheDatenWK81d sFeiertag, sDatum
        Case Is = 3
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 8     'drucken
        
'            drucke_Behandlungstexte
            
        
        Case Is = 6     'Beenden
            Unload frmWK81d
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "DATUM"
                SpaltennummerDatum = i
            Case Is = "FEIERTAGBEZ"
                SpaltennummerFeiertag = i
            Case Is = "ANWENDEN"
                SpaltennummerANWENDEN = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    WKL81dPositionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If NewTableSuchenDBKombi("FEIERTAGE", gdBase) = False Then
        CreateTableT2 "FEIERTAGE", gdBase
    
        'bundesweit fix
        insert_Feiertag "01.01.", "", "Neujahr", 1, 1
        insert_Feiertag "01.05.", "", "Tag der Arbeit", 1, 1
        insert_Feiertag "03.10.", "", "Tag der Deutschen Einheit", 1, 1
        insert_Feiertag "25.12.", "", "1. Weihnachtstag", 1, 1
        insert_Feiertag "26.12.", "", "2. Weihnachtstag", 1, 1
        
        'bundesweit beweglich
        insert_Feiertag "18.04.", "2014", "Karfreitag", 1, 1
        insert_Feiertag "04.04.", "2015", "Karfreitag", 1, 1
        insert_Feiertag "25.03.", "2016", "Karfreitag", 1, 1
        
        
        insert_Feiertag "21.04.", "2014", "Ostermontag", 1, 1
        insert_Feiertag "06.04.", "2015", "Ostermontag", 1, 1
        insert_Feiertag "28.03.", "2016", "Ostermontag", 1, 1
        
        insert_Feiertag "29.05.", "2014", "Christi Himmelfahrt", 1, 1
        insert_Feiertag "14.05.", "2015", "Christi Himmelfahrt", 1, 1
        insert_Feiertag "05.05.", "2016", "Christi Himmelfahrt", 1, 1
        
        insert_Feiertag "09.06.", "2014", "Pfingstmontag", 1, 1
        insert_Feiertag "25.05.", "2015", "Pfingstmontag", 1, 1
        insert_Feiertag "16.05.", "2016", "Pfingstmontag", 1, 1
        
        
        'Nicht bundesweit fix
        insert_Feiertag "06.01.", "", "Heilige Drei Könige", 0, 0
        insert_Feiertag "15.08.", "", "Mariä Himmelfahrt", 0, 0
        insert_Feiertag "31.10.", "", "Reformationstag", 0, 0
        insert_Feiertag "01.11.", "", "Allerheiligen", 0, 0
        
        'Nicht bundesweit beweglich
        insert_Feiertag "19.06.", "2014", "Fronleichnam", 0, 0
        insert_Feiertag "04.06.", "2015", "Fronleichnam", 0, 0
        insert_Feiertag "26.05.", "2016", "Fronleichnam", 0, 0
        
        insert_Feiertag "20.11.", "2013", "Buß- und Bettag", 0, 0
        insert_Feiertag "19.11.", "2014", "Buß- und Bettag", 0, 0
        insert_Feiertag "18.11.", "2015", "Buß- und Bettag", 0, 0
        insert_Feiertag "16.11.", "2016", "Buß- und Bettag", 0, 0
    
    End If

    
    Zeige_Daten_81d
    LeereFelderWK81d
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub WKL81dPositionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 120
    Frame1.Height = 6735
    Frame1.Width = 11655
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL81dPositionieren"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "FEIERTAGE_TEMP", gdBase
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
Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR

    Dim sFeiertag As String
    Dim sDatum  As String

    If MSHFLEX1.Row = 1 Then
        sortierenHGrid MSHFLEX1
    Else
        If MSHFLEX1.Col = SpaltennummerANWENDEN Then
            Select Case MSHFLEX1.Text()
                Case "nicht anwenden"
                    MSHFLEX1.Text = "anwenden"
                    MSHFLEX1.CellFontBold = True
                    MSHFLEX1.CellForeColor = vbGreen
                    
                    sFeiertag = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerFeiertag))
                    sDatum = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerDatum))

            
                    uPdate_Feiertage sFeiertag, sDatum, "anwenden"

                    
                Case "anwenden"
                    MSHFLEX1.Text = "nicht anwenden"
                    MSHFLEX1.CellFontBold = True
                    MSHFLEX1.CellForeColor = vbRed
                    
                    sFeiertag = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerFeiertag))
                    sDatum = Trim(MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerDatum))

            
                    uPdate_Feiertage sFeiertag, sDatum, "nicht anwenden"
            End Select
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub uPdate_Feiertage(sFeiert As String, sdat As String, sAnwendungsart As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Update Feiertage set  "
    If sAnwendungsart = "anwenden" Then
        cSQL = cSQL & " Anwenden = 1 "
    ElseIf sAnwendungsart = "nicht anwenden" Then
        cSQL = cSQL & " Anwenden = 0 "
    End If
    cSQL = cSQL & " where FEIERTAGBEZ = '" & sFeiert & "' "
    cSQL = cSQL & " and FDAT = '" & Left(sdat, 6) & "'"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "uPdate_Feiertage"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command1_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Vorgaben Feiertage ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



