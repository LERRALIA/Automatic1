VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmWKL184 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Rewe Warengruppen"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   5
      Left            =   9480
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
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
      Caption         =   "alle auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   3
      Left            =   9480
      TabIndex        =   0
      Top             =   7800
      Width           =   2175
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
      Caption         =   "Weiter"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   18
      FixedCols       =   2
      ForeColorSel    =   8454143
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   0
      Left            =   11280
      TabIndex        =   6
      Top             =   240
      Width           =   405
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
      ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
      ToolTipTitle    =   "Spaltenanordung"
      ButtonStyle     =   2
      Caption         =   ""
      Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
      Picture         =   "frmWKL184.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Markieren Sie die gewünschten Sortimente (Doppelklick)"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   9015
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Rewe - Welche Sortimente möchten Sie übernehmen?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL184"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerREWEWGRTEXT As Byte
Dim SpaltennummerWAHL As Byte


Private Sub flex(krit As String)
On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = SpaltennummerWAHL
        MSFlexGrid1.Row = lcount

        Select Case krit
            Case "auswählen"
                MSFlexGrid1.Text = "ausgewählt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbGreen
                
            Case "entfernen"
                MSFlexGrid1.Text = "nicht ausgewählt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbRed
        End Select
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "flex"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub speicherDieWahl()
On Error GoTo LOKAL_ERROR
    
    Dim lcount              As Long
    Dim sSQL                As String
    Dim sReweWGRTEXT        As String

    anzeige "normal", "", lblanzeige
    Screen.MousePointer = 11
    
    MSFlexGrid1.Redraw = False
    
    For lcount = 1 To MSFlexGrid1.Rows - 1
    
        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = SpaltennummerREWEWGRTEXT
        sReweWGRTEXT = MSFlexGrid1.Text
        
        MSFlexGrid1.Col = SpaltennummerWAHL
        
        Select Case MSFlexGrid1.Text()
            Case "nicht ausgewählt"
                sSQL = "Update ReweWGR set aktgew = false "
                sSQL = sSQL & " where REWEWGRTEXT = '" & sReweWGRTEXT & "'"
                gdBase.Execute sSQL, dbFailOnError
            Case "ausgewählt"
            
                sSQL = "Update ReweWGR set aktgew = true "
                sSQL = sSQL & " where REWEWGRTEXT = '" & sReweWGRTEXT & "'"
                gdBase.Execute sSQL, dbFailOnError
        End Select
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
    Screen.MousePointer = 0
    Unload frmWKL184
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDieWahl"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        Case 0
            gsZSpalte = "REWEWGRTEXT"
            gstab = "REWEWGR"
            frmWKL36.Show 1
            'fertig
            
            ZeigeRewegruppe
            If MSFlexGrid1.Visible = True Then
                MSFlexGrid1.Col = 1
                MSFlexGrid1.Row = 2
                MSFlexGrid1.SetFocus
            End If
        
        Case 3
            speicherDieWahl
            
        Case 5
            If Command4(5).Caption = "alle auswählen" Then
                Command4(5).Caption = "alle entfernen"
                flex "auswählen"
            
            Else
                Command4(5).Caption = "alle auswählen"
                flex "entfernen"
            End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    anzeige "normal", "", lblanzeige
    
    ZeigeRewegruppe
    If MSFlexGrid1.Visible = True Then
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = 2
        MSFlexGrid1.SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "REWEWGRTEXT"
                SpaltennummerREWEWGRTEXT = i
            Case Is = "AKTGEW"
                SpaltennummerWAHL = i
        End Select
    Next i
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex164()
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
    Dim cSQL        As String
   
    cSQL = "Select * from REWEWGR order by REWEWGRTEXT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                            
                            Case Is = "auswählen"
                                
                                .Row = lrow
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    If rsrs(sSpaltenbez(i)) = True Then
                                        .Text = "ausgewählt"
                                        .CellFontBold = True
                                        .CellForeColor = vbGreen
                                        
                                    Else
                                        .Text = "nicht ausgewählt"
                                        .CellFontBold = True
                                        .CellForeColor = vbRed
                                    End If
                                Else
                                    .Text = "ausgewählt"
                                    .CellFontBold = True
                                    .CellForeColor = vbRed
                                End If
                                
                                
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        
        rsrs.Close
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex164"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeRewegruppe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    loeschNEW "REWEWGR", gdBase
    CreateTableT2 "REWEWGR", gdBase
    
    sSQL = " Insert into REWEWGR select count(libesnr) as anzahl, MARKE as REWEWGRTEXT  from IMPORTPRI "
    sSQL = sSQL & " group by MARKE "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update REWEWGR set aktgew = false "
    gdBase.Execute sSQL, dbFailOnError
        
    ZeigeReweTAB
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeRewegruppe"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeReweTAB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    Set recAnz = gdBase.OpenRecordset("REWEWGR")
    If recAnz.EOF Then
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        
        anzeige "rot", "Keine Warengruppen gefunden!", lblanzeige
        recAnz.Close
        Exit Sub
    End If
    recAnz.Close
    
    Screen.MousePointer = 11

    Tabcheck "REWEWGR"
    
    FormatGridOverTablay "REWEWGR"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex164
        ermittlespalten
        
        .Redraw = False
    
        Tabellenbreiteanpassen MSFlexGrid1, 1.85 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeReweTAB"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
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
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR

    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    Else
        MSFlexGrid1.Col = SpaltennummerWAHL
        Select Case MSFlexGrid1.Text()
            Case "nicht ausgewählt"
                MSFlexGrid1.Text = "ausgewählt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbGreen
            Case "ausgewählt"
                MSFlexGrid1.Text = "nicht ausgewählt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbRed
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Rewe Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


