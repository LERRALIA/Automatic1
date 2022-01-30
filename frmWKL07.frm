VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL07 
   Caption         =   "Produktgruppenbearbeitung"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   8175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1440
      Width           =   6975
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   8175
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Caption         =   "Felder leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   4
      Top             =   6120
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Caption         =   "Neue PGN"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   2
      Top             =   7320
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      Index           =   2
      Left            =   9600
      TabIndex        =   0
      Top             =   7920
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      TabIndex        =   13
      Top             =   8040
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "PGN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Produktgruppen-Bezeichnung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Produktgruppenbearbeitung"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmWKL07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function pruef() As Boolean
    On Error GoTo LOKAL_ERROR
    
    pruef = False
    
    If Text1(0).Text = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(Text1(0).Text) Then
        Exit Function
    End If
    
    If Text1(1).Text = "" Then
                
        Exit Function
    Else
        Text1(1).Text = SwapStr(Text1(1).Text, "'", " ")
        Text1(1).Text = SwapStr(Text1(1).Text, "*", " ")
        If Text1(1).Text = "" Then
                
            Exit Function
        End If
    End If
        
    pruef = True
            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pruef"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
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
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Select Case Index
        
        Case Is = 0     'Speichern
        
            If pruef Then
                If SchreibeDatenWKL12 Then
                    FuelleListbox2WKL12
                    InitDialogWKL12
                End If
            Else
                anzeigeNew "rot", "Bitte überprüfen Sie Ihre Eingaben!", lblanzeige
            End If
        Case Is = 1     'Leeren
            InitDialogWKL12
            Text1(0).SetFocus
        Case Is = 2     'Beenden
            Unload frmWKL07
        Case Is = 4     'Neue AGN
            InitDialogWKL12
            Text1(0).SetFocus
        Case Is = 5    'Drucken
            reportbildschirm "dWKL12a", "aWKL05"
        Case Is = 6
            If List2.ListIndex < 0 Then
                anzeigeNew "rot", "Bitte einen Eintrag in der Liste markieren!", lblanzeige
                List2.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            MoveList2FelderWKL12
            LoescheAGN
            InitDialogWKL12
        
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    WKL05Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    InitDialogWKL12
    List1.Clear
    List2.Clear
    List1.AddItem "PGN    Produktgruppenbezeichnung"
    
    FuelleListbox2WKL12
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheAGN()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cFeld As String
    
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    
    cSQL = "Delete from PGNDBF where PGN = " & cFeld & " "
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    FuelleListbox2WKL12
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheAGN"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MoveList2FelderWKL12()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cFeld As String
    
    cLBSatz = List2.list(List2.ListIndex)
    
    cFeld = Mid(cLBSatz, 1, 3)
    cFeld = Trim$(cFeld)
    Text1(0).Text = cFeld
    
    cFeld = Mid(cLBSatz, 5, 35)
    cFeld = Trim$(cFeld)
    Text1(1).Text = cFeld
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveList2FelderWKL12"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SchreibeDatenWKL12() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim rs1     As Recordset
    Dim cFeld   As String
    Dim dFeld   As Double
    Dim cAGTE   As String
    
    SchreibeDatenWKL12 = False
    
    cFeld = Trim$(Text1(0).Text)
    cAGTE = Trim(Text1(1).Text)
    
    cSQL = "Select * from PGNDBF where PGN = " & cFeld & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
        cSQL = "Select PGNBEZEICH from PGNDBF where Ucase(PGNBEZEICH) = '" & UCase(cAGTE) & "' "
        Set rs1 = gdBase.OpenRecordset(cSQL)
        If Not rs1.EOF Then
            anzeigeNew "rot", "Diese Produktgruppenbezeichnung ist schon vergeben.", lblanzeige
            Text1(1).SetFocus
            rs1.Close: Set rs1 = Nothing
            Exit Function
        End If
        rs1.Close: Set rs1 = Nothing
    
    End If
    
    cFeld = Text1(0).Text
    rsrs!PGN = cFeld
    rsrs!PGNBEZEICH = cAGTE
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    SchreibeDatenWKL12 = True
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL12"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub FuelleListbox2WKL12()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim lAnz        As Long
    
    List2.Clear
    
    cSQL = "Select * from PGNDBF order by PGN "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!PGN) Then
                dWert = rsrs!PGN
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "#####0")
            cLBSatz = Space(3 - Len(cFeld)) & cFeld & "    "
            
            If Not IsNull(rsrs!PGNBEZEICH) Then
                cFeld = rsrs!PGNBEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld

            
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lAnz = 0 Then
        anzeigeNew "rot", "Keine Produktgruppen wurden ermittelt.", lblanzeige
    ElseIf lAnz = 1 Then
        anzeigeNew "normal", lAnz & " Produktgruppe wurde ermittelt.", lblanzeige
    Else
        anzeigeNew "normal", lAnz & " Produktgruppen wurden ermittelt.", lblanzeige
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListbox2WKL12"
    Fehler.gsFehlertext = "Im Programmteil Artikelgruppen bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL05Positionieren()
    On Error GoTo LOKAL_ERROR
    
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL05Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    Select Case Index
        Case 0
            cValid = "1234567890" & Chr$(8)
        Case 1
            cValid = Chr$(KeyAscii)
    End Select
    
    If InStr(cValid, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(Chr$(KeyAscii))
    End If
        

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Produktgruppen bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
