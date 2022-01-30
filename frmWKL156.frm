VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL156 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Markenbearbeitung"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   17
      Top             =   360
      Width           =   345
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   35
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   7800
      TabIndex        =   11
      Top             =   1080
      Width           =   3855
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2520
         Style           =   2  'Dropdown-Liste
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Übersicht"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Vormonatsauswertung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   2895
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1455
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   5
      TabIndex        =   7
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   6
      Top             =   6720
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   9840
      TabIndex        =   1
      Top             =   8040
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Suche"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   8040
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kürzel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Markenbezeichnung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Markenbearbeitung"
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
      Width           =   7575
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
End
Attribute VB_Name = "frmWKL156"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "MAVM" & srechnertab, gdBase
    loeschNEW "MAVMPRINT", gdBase
    loeschNEW "MAVMKOPF", gdBase
    
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
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim iMonat As Byte
    Dim iJahr As Integer
    
    
    
    Screen.MousePointer = 11
    Select Case Index
        
        Case 0     'Speichern
        
            SchreibeDatenWKL12
            FuelleListbox2WKL12 Text1(2).Text
            InitDialogWKL12

        Case 2      'Beenden
            Unload frmWKL156
        
        Case 6 'Marke
            If Option2(0).Value = True Then
            
                
                Uebersicht_erstellen
                reportbildschirm "", "aZEN81"
            ElseIf Option2(1).Value = True Then
                iMonat = CByte(Mid$(Combo3.Text, 1, InStr(1, Combo3.Text, "/") - 1))
                iJahr = CInt(Right(Combo3.Text, 4))
                            
                VorMonatsAuswertungMarke iMonat, iJahr
            End If
        Case 11
            gsHelpstring = "Markenbearbeitung"
            frmWKL110.Show 1
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Uebersicht_erstellen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    loeschNEW "LINBEZPRINT", gdBase
    CreateTableT2 "LINBEZPRINT", gdBase
    
    sSQL = "Insert into LINBEZPRINT Select * from LINBEZ"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LINBEZPRINT where Kuerzel = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LINBEZPRINT where Kuerzel is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LINBEZPRINT inner join LISRT on LINBEZPRINT.linr = LISRT.linr "
    sSQL = sSQL & " set LINBEZPRINT.liefbez = lisrt.liefbez"
    gdBase.Execute sSQL, dbFailOnError
    
    
    Dim lLinr As Long
    Dim lLpz As Long
    Dim lLagerST As Long
    Dim rsrs As Recordset
    
    cSQL = "Select * from LINBEZPRINT "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF


            If Not IsNull(rsrs!LPZ) Then
                lLpz = rsrs!LPZ

                If Not IsNull(rsrs!linr) Then
                    lLinr = rsrs!linr

                
                    lLagerST = LAGERStückErmittlungJetztLPZ(lLinr, lLpz)
                End If
                
            Else
                lLagerST = 0
            End If

            rsrs.Edit
            rsrs!BESTAND = lLagerST
            rsrs.Update

        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
'    sSQL = "Update LINBEZPRINT inner join Artikel on LINBEZPRINT.linr = LISRT.linr "
'    sSQL = sSQL & " set LINBEZPRINT.bestand = Artikel.bestand"
'    gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebersicht_erstellen"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub VorMonatsAuswertungMarke(imon As Byte, iJahr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsrs1 As Recordset
    Dim cMarke As String
    Dim lLinr As Long
    Dim lLpz As Long
    Dim lAnz As Long
    Dim lAnz2 As Long
    Dim j As Integer
    
    Dim lLagerST As Long

    Dim lAbsatzVM As Long
    Dim dUmsatzVM As Double
    Dim dUmsatzVVM As Double
    Dim lEinkaufVM As Long
    
    Dim lAbsatzaktJ As Long
    Dim dUmsatzaktJ As Double
    Dim dUmsatzVaktJ As Double
    Dim lEinkaufaktJ As Long
    
    Dim dUmsatzabs As Double
    Dim dUmsatzrela As Double
    
    Dim dSummeUmsatzVM As Double
    Dim dSummeUmsatzAKTJ As Double
    
    anzeige "normal", "Daten werden ermittelt...", lblAnzeige
    
    loeschNEW "MAVM" & srechnertab, gdBase
    CreateTableT2 "MAVM" & srechnertab, gdBase
   
    cSQL = "Insert Into MAVM" & srechnertab & " Select   "
    cSQL = cSQL & " distinct(marke) as MARKENBEZ  "
    cSQL = cSQL & " from LINBEZ where "
    cSQL = cSQL & " Linr < 500000 "
'    cSQL = cSQL & " Linr = 312130 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "mit Detailzahlen: es werden Hintergrunddaten zusammengefasst...", lblAnzeige

    If UMS_LPZaktuell = False Then
        ErzeugeLpzUmsatz
    End If

    anzeige "normal", "Lagerwerte werden ermittelt...", lblAnzeige

    LagerwerteschreibenLPZJetzt lblAnzeige, 0

    cSQL = "Select * from MAVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lAnz = lAnz - 1
'            anzeige "normal", "Marke: " & cMarke & " noch " & CStr(lAnz) & " Marken ...", lblanzeige

            If Not IsNull(rsrs!Markenbez) Then
                cMarke = Trim(rsrs!Markenbez)
                cMarke = SwapStr(cMarke, "'", "")
                
                lAbsatzVM = 0
                dUmsatzVM = 0
                dUmsatzVVM = 0
                lEinkaufVM = 0
                lLagerST = 0
                
                lAbsatzaktJ = 0
                dUmsatzaktJ = 0
                dUmsatzVaktJ = 0
                lEinkaufaktJ = 0
                
                cSQL = "Select LINR,lpz from linbez where marke = '" & cMarke & "'"
                Set rsrs1 = gdBase.OpenRecordset(cSQL)
                If Not rsrs1.EOF Then
                    rsrs1.MoveLast
                    lAnz2 = rsrs1.RecordCount
                    rsrs1.MoveFirst
                    Do While Not rsrs1.EOF
                    
                        lAnz2 = lAnz2 - 1
                        anzeige "normal", "noch " & CStr(lAnz) & " Marken ..." & CStr(lAnz2) & " " & cMarke, lblAnzeige
                    
                        If Not IsNull(rsrs1!linr) Then
                            lLinr = rsrs1!linr
                        Else
                            lLinr = 0
                        End If
                        
                        If Not IsNull(rsrs1!LPZ) Then
                            lLpz = rsrs1!LPZ
                        Else
                            lLpz = 0
                        End If

                        lAbsatzVM = lAbsatzVM + ermgesAbsatzLPZ(imon, iJahr, lLinr, lLpz)
                        dUmsatzVM = dUmsatzVM + ermgesUmsatzLpz(imon, iJahr, lLinr, lLpz)
                        dUmsatzVVM = dUmsatzVVM + ermgesUmsatzLpz(imon, iJahr - 1, lLinr, lLpz)
                        lEinkaufVM = lEinkaufVM + EinkaufsStückermittlungLPZ(CStr(lLinr), gdBase, iJahr, imon, lLpz)
        
                        If imon = 12 Then
                            lAbsatzaktJ = lAbsatzaktJ + ermgesAbsatzLPZ(0, iJahr, lLinr, lLpz)
                            dUmsatzaktJ = dUmsatzaktJ + ermgesUmsatzLpz(0, iJahr, lLinr, lLpz)
                            dUmsatzVaktJ = dUmsatzVaktJ + ermgesUmsatzLpz(0, iJahr - 1, lLinr, lLpz)
                            lEinkaufaktJ = lEinkaufaktJ + EinkaufsStückermittlungLPZ(CStr(lLinr), gdBase, iJahr, 0, lLpz)
                        Else
                            For j = 1 To imon
                                lAbsatzaktJ = lAbsatzaktJ + ermgesAbsatzLPZ(CByte(j), iJahr, lLinr, lLpz)
                                dUmsatzaktJ = dUmsatzaktJ + ermgesUmsatzLpz(CByte(j), iJahr, lLinr, lLpz)
                                dUmsatzVaktJ = dUmsatzVaktJ + ermgesUmsatzLpz(CByte(j), iJahr - 1, lLinr, lLpz)
                                lEinkaufaktJ = lEinkaufaktJ + EinkaufsStückermittlungLPZ(CStr(lLinr), gdBase, iJahr, CByte(j), lLpz)
                            Next j
                        End If
        
                        lLagerST = lLagerST + LAGERStückErmittlungJetztLPZ(lLinr, lLpz)
                    
                    rsrs1.MoveNext
                    Loop
                    
                End If
                rsrs1.Close
            Else
                lLagerST = 0
                lAbsatzVM = 0
                dUmsatzVM = 0
                dUmsatzVVM = 0

                lEinkaufVM = 0

                lAbsatzaktJ = 0
                dUmsatzaktJ = 0
                dUmsatzVaktJ = 0
                lEinkaufaktJ = 0

            End If

            rsrs.Edit
            rsrs!LAGERST = lLagerST

            rsrs!ABSATZVM = lAbsatzVM
            rsrs!UmsatzVM = dUmsatzVM
            rsrs!UmsatzVVM = dUmsatzVVM
            rsrs!EINKAUFVM = lEinkaufVM

            rsrs!AbsatzaktJ = lAbsatzaktJ
            rsrs!EINKAUFaktJ = lEinkaufaktJ
            rsrs!Umsatzaktj = dUmsatzaktJ
            rsrs!UmsatzVAKTJ = dUmsatzVaktJ

            dUmsatzabs = 0
            dUmsatzabs = dUmsatzVM - dUmsatzVVM
            dUmsatzrela = 0
            If dUmsatzVM <> 0 Then
                dUmsatzrela = Round(100 * dUmsatzabs / dUmsatzVM, 0)
            End If

            rsrs!UMSATZMRELA = dUmsatzrela

            dUmsatzabs = 0
            dUmsatzabs = dUmsatzaktJ - dUmsatzVaktJ
            dUmsatzrela = 0
            If dUmsatzaktJ <> 0 Then
                dUmsatzrela = Round(100 * dUmsatzabs / dUmsatzaktJ, 0)
            End If

            rsrs!UMSATZJRELA = dUmsatzrela

            If lAbsatzVM <> 0 Then
                rsrs!VKPREISPROSTCK = dUmsatzVM / lAbsatzVM
            End If

            If lAbsatzVM <> 0 Then
                rsrs!LagerRWM = lLagerST / lAbsatzVM
            End If

            rsrs.Update

        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    dSummeUmsatzVM = 0
    cSQL = "Select sum(UmsatzVM) as maxi from MAVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSummeUmsatzVM = rsrs!maxi
        End If
    End If
    rsrs.Close

    dSummeUmsatzAKTJ = 0
    cSQL = "Select sum(Umsatzaktj) as maxi from MAVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSummeUmsatzAKTJ = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = "Select * from MAVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast

        rsrs.MoveFirst
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr

            End If

            rsrs.Edit

            If dSummeUmsatzVM <> 0 Then
                rsrs!MarktanteilM = 100 * rsrs!UmsatzVM / dSummeUmsatzVM
            End If

            If dSummeUmsatzAKTJ <> 0 Then
                rsrs!MarktanteilJ = 100 * rsrs!Umsatzaktj / dSummeUmsatzAKTJ
            End If

            rsrs.Update

        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    loeschNEW "MATT", gdBase

    cSQL = "Select * into MATT from MAVM" & srechnertab
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "MAVM" & srechnertab, gdBase
    CreateTableT2 "MAVM" & srechnertab, gdBase

    cSQL = "Insert Into MAVM" & srechnertab & " Select Top 50 UmsatzVM,* "
    cSQL = cSQL & " from MATT "
    cSQL = cSQL & " where UmsatzVM > 0 order by UmsatzVM desc "
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "MATT", gdBase

    'Platzierungen

    lAnz = 1
    cSQL = "Select * from MAVM" & srechnertab & " order by Umsatzvm desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzUmsatzM = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    lAnz = 1
    cSQL = "Select * from MAVM" & srechnertab & " order by AbsatzVM desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzSTCKM = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    lAnz = 1
    cSQL = "Select * from MAVM" & srechnertab & " order by Umsatzaktj desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzUmsatzJ = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    lAnz = 1
    cSQL = "Select * from MAVM" & srechnertab & " order by Absatzaktj desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzSTCKJ = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    loeschNEW "MAVMPRINT", gdBase
    CreateTableT2 "MAVMPRINT", gdBase

    cSQL = "Insert into MAVMPRINT select * from MAVM" & srechnertab
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "MAVM" & srechnertab, gdBase

    'Kopfdaten
    loeschNEW "MAVMKOPF", gdBase
    CreateTableT2 "MAVMKOPF", gdBase

    Dim sdat As String
    Dim sBasis As String
    sdat = MonthName(imon) & " " & iJahr
    sBasis = "1 Geschäft"

    cSQL = "Insert into MAVMKOPF (UEBER,Auswertungsdat,Basis) values ('Marken/Depots','" & sdat & "','" & sBasis & "')"
    gdBase.Execute cSQL, dbFailOnError

    anzeige "normal", "Druckvorschau wird erstellt...", lblAnzeige
    reportbildschirm "", "zZEN17d"
    
    anzeige "normal", "", lblAnzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VorMonatsAuswertungMarke"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    List1.Clear
    List2.Clear
    List1.AddItem "Kürzel   Markenbezeichnung"
    
    FuelleListbox2WKL12 Text1(2).Text
    
    fuellecombo
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim iMonat As Integer
    Dim iJahr As Integer
    
    iMonat = Month(Now)
    iJahr = Year(Now)
    
    With Combo3
        .Clear
        For i = 1 To 12
        
            If iMonat = 1 Then
                iMonat = 12
                iJahr = iJahr - 1
            Else
                iMonat = iMonat - 1
                iJahr = iJahr
            End If
            
            .AddItem iMonat & "/" & iJahr
            If .Text = "" Then
                .Text = iMonat & "/" & iJahr
            End If
            
        Next i
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MoveList2FelderWKL12()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cFeld As String
    
    cLBSatz = List2.list(List2.ListIndex)
    
    cFeld = Mid(cLBSatz, 1, 5)
    cFeld = Trim$(cFeld)
    Text1(0).Text = cFeld
    
    cFeld = Mid(cLBSatz, 7, 35)
    cFeld = Trim$(cFeld)
    Text1(1).Text = cFeld
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveList2FelderWKL12"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWKL12()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cFeld   As String
    Dim cFeld2   As String
    
    cFeld = UCase(Trim$(Text1(0).Text))
    cFeld2 = Trim$(Text1(1).Text)
    
    cSQL = "Update LINBEZ set kuerzel = ''"
    cSQL = cSQL & " where MARKE = '" & cFeld2 & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update LINBEZ set kuerzel = '" & cFeld & "'"
    cSQL = cSQL & " where MARKE = '" & cFeld2 & "' "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL12"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuelleListbox2WKL12(cSuch As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim lAnz        As Long
    
    List2.Clear
    
    cSQL = "Select distinct(marke),KUERZEL from LINBEZ where not lpz = 0 and not Marke is null and not Marke = ''  "
    
    If cSuch <> "" Then
        cSQL = cSQL & " and KUERZEL Like '" & cSuch & "*' "
    End If
    
    cSQL = cSQL & " order by marke"
    
    lAnz = 0
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kuerzel) Then
                cFeld = rsrs!Kuerzel
            Else
                cFeld = ""
            End If
            
            cLBSatz = cFeld & Space(9 - Len(cFeld))
            
            If Not IsNull(rsrs!MARKE) Then
                cFeld = rsrs!MARKE
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld

            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If lAnz = 0 Then
        cSQL = "Select distinct(marke),KUERZEL from LINBEZ where not lpz = 0 and not Marke is null and not Marke = ''  "
        
        If cSuch <> "" Then
            cSQL = cSQL & " and marke Like '*" & cSuch & "*' "
        End If
        
        
        cSQL = cSQL & " order by marke"
        
        lAnz = 0
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveLast
            lAnz = rsrs.RecordCount
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!Kuerzel) Then
                    cFeld = rsrs!Kuerzel
                Else
                    cFeld = ""
                End If
                
                cLBSatz = cFeld & Space(9 - Len(cFeld))
                
                If Not IsNull(rsrs!MARKE) Then
                    cFeld = rsrs!MARKE
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cLBSatz = cLBSatz & cFeld
    
                List2.AddItem cLBSatz
                
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close
        
    End If
    
    If lAnz = 0 Then
        anzeige "rot2", "Es wurden keine Marken ermittelt.", lblAnzeige
    ElseIf lAnz = 1 Then
        anzeige "normal", lAnz & " Marken wurde ermittelt.", lblAnzeige
    Else
        anzeige "normal", lAnz & " Marken wurden ermittelt.", lblAnzeige
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListbox2WKL12"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InitDialogWKL12()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InitDialogWKL12"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_Click()
On Error GoTo LOKAL_ERROR
    
    MoveList2FelderWKL12
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 2 Then
        FuelleListbox2WKL12 Text1(2).Text
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim llen As Long

    With Text1(Index)
    
        .BackColor = glSelBack1
        
        llen = Len(.Text)
        .SelStart = llen   ' Textende markieren, damit es bei neuer Eingabe gleich wieder gelöscht wird
        .SelLength = Len(.Text) - llen
        
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
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
        Case 0
            cValid = gcUPPER & gcLower & Chr$(8)
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
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Marken bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
